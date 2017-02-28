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
    class Program
    {
        public static string orahlperparam;
        static void Main(string[] args)
        {
            orahlperparam = (args.Count() > 0 ? args : new string[]{"-p Audit.exe"})[0].Split()[0];
            StatisticHelper sh = new StatisticHelper(Program.orahlperparam);
            DataTableHelper dth = new DataTableHelper();
      //      sh.PrepareTabs();
      //      sh.ShowDT();
            

            int[] row = {2,4},
                c单位 = {1,2},
                c仪器 = {2,5},
            c增加 = {3,7},
            c事件数 = {4,9},
            c审核率 = {5,11},
            c月报 = {6,12},
            c质量 = {7,10},
            c总 = {8,15,16};
          

            DataTable[] montabs = new DataTable[13];
            DataTable atab = new DataTable();
            excel.Application eapp = new excel.Application();
            
            int mon;
            for (mon = 1; mon <= 12; mon++)
            {
                string fname = string.Format("c:\\15\\附件：2016年{0}月全国各区域前兆台网数据跟踪评分详表.xlsx", mon);
                excel.Workbook book = eapp.Workbooks.Open(fname);
                excel.Worksheet sheet = book.Sheets["月评得分总表"];
                DataTable mt = new DataTable();
                mt.Columns.Add("单位");
                mt.Columns.Add("仪器", typeof(double));
                mt.Columns.Add("增加", typeof(double));
                mt.Columns.Add("事件数", typeof(double));
                mt.Columns.Add("审核率", typeof(double));
                mt.Columns.Add("月报", typeof(double));
                mt.Columns.Add("质量", typeof(double));
                mt.Columns.Add("总", typeof(double));
                if (mon == 1)
                {
                    atab = mt.Clone();
                    atab.PrimaryKey = new DataColumn[] { atab.Columns["单位"] };
                    foreach (string au in GDef.abbrunitnamelist)
                    {
                        if (au == "震防中心")
                            continue;
                        atab.Rows.Add(au, 0, 0, 0, 0, 0, 0, 0);
                    }
                }
                int j, jz;
                if (mon < 6)
                {
                    j = 0;
                    jz = 0;
                }
                else
                {
                    j = 1;
                    jz = 1;
                    if (mon >= 7)
                        jz = 2;
                }
                object[,] t = sheet.Range[sheet.Cells[row[j], 1], sheet.Cells[row[j] + 50, c总[jz]]].Value;
                book.Close(false);
                for (int i = 1; i <= t.GetLength(0); i++)
                {
                    bool find = false;
                    foreach(string au in GDef.abbrunitnamelist)
                    {
                        if (t[i, c单位[j]] == null)
                            break;
                        if (t[i, c单位[j]].ToString() == au)
                        {
                            find = true;
                            break;
                        }
                    }
                    if (find == false)
                        break;
                    if (t[i, c单位[j]].ToString() == "震防中心")
                        continue;
                    DataRow dr = mt.NewRow();
                    dr["单位"] = t[i, c单位[j]];
                    DataRow adr = atab.Rows.Find(dr["单位"]);
                    dr["仪器"] = t[i, c仪器[j]];
                    adr["仪器"] = adr.Field<double>("仪器") + dr.Field<double>("仪器");
                    dr["增加"] = t[i, c增加[j]];
                    adr["增加"] = adr.Field<double>("增加") + dr.Field<double>("增加");
                    dr["事件数"] = t[i, c事件数[j]];
                    adr["事件数"] = adr.Field<double>("事件数") + dr.Field<double>("事件数");
                    dr["审核率"] = t[i, c审核率[j]];
                    adr["审核率"] = adr.Field<double>("审核率") + dr.Field<double>("审核率");
                    if (mon < 6)
                    {
                        dr["月报"] = t[i, c月报[j]];
                    }
                    else
                    {
                        dr["月报"] = Convert.ToDouble(t[i, c月报[j]]) + Convert.ToDouble(t[i, c月报[j] + 1]) + Convert.ToDouble(t[i, c月报[j] + 2]);
                    }
                    adr["月报"] = adr.Field<double>("月报") + dr.Field<double>("月报");
                    dr["质量"] = t[i, c质量[j]]?? 0;
                    adr["质量"] = adr.Field<double>("质量") + dr.Field<double>("质量");
                    dr["总"] = t[i, c总[jz]];
                    adr["总"] = adr.Field<double>("总") + dr.Field<double>("总");
                    mt.Rows.Add(dr);
                }
                montabs[mon] = mt;
            }
            foreach (DataRow r in atab.Rows)
            {
                foreach (DataColumn c in atab.Columns)
                {
                    if (c.ColumnName == "单位")
                        continue;
                    r[c] = Math.Round(r.Field<double>(c) / 12.0, 2);
                }
            }
            excel.Workbook book2 = eapp.Workbooks.Add();
            atab.DefaultView.Sort = "总 desc";
            dth.DTToExcelSheet(atab.DefaultView.ToTable(), book2, null, book2.Worksheets[1]);
            StreamWriter sw = new StreamWriter("c:\\qztemp\\12output1.txt");

            foreach (DataColumn c in atab.Columns)
            {
                atab.DefaultView.Sort = c.ColumnName + " desc";
                DataTable dt = atab.DefaultView.ToTable(false, "单位");
                sw.Write("按" + c.ColumnName + "分从高至低 ");
                foreach (DataRow r in dt.Rows)
                {
                    sw.Write(r["单位"] + "、");
                }
                sw.WriteLine();
            }


            sw.Close();

            eapp.Visible = true;

        //    eapp.Quit();
        }
    }
}
