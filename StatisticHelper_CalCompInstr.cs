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
        double CalCompInstr(DataTable logtab, string instrid, DateTime begin_contain, DateTime end_contain, DateMergerPrecision p, string instrid_colname = "sp", string startdate_colname = "start_date", string enddate_colname = "end_date")
        {
            DataView dv = new DataView(logtab);
            dv.RowFilter = string.Format("{0} = '{1}'", instrid_colname, instrid);
            if (dv.Count == 0)
            {
                return 0;
            }
            DataTable instrlog = dv.ToTable();
            DateMerger dm = new DateMerger(begin_contain, end_contain, p);
            foreach (DataRow r in instrlog.Rows)
            {
                dm.Merge(new DatePair(Convert.ToDateTime(r[startdate_colname]), Convert.ToDateTime(r[enddate_colname])));
            }
            return dm.CalComp();
        }
    }
}