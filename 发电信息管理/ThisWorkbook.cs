using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml.Linq;
using Microsoft.Office.Tools.Excel;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace 发电信息管理
{
    public partial class ThisWorkbook
    {
        public System.Data.DataTable baseTable = new System.Data.DataTable("baseTable");
        public System.Data.DataTable recordTable = new System.Data.DataTable("recordTable");



        private void ThisWorkbook_Startup(object sender, System.EventArgs e)
        {
            InitBaseTable();
            InitRecordTable();
           
            Worksheets["输入数据"].Activate();
        }

        private void ThisWorkbook_Shutdown(object sender, System.EventArgs e)
        {
        }

    
        //生成基础数据表
        private void InitBaseTable()
        {


            Excel.Worksheet baseSheet = new Excel.Worksheet();
            baseSheet = Globals.ThisWorkbook.Worksheets["基础信息"];

            //将数据表第一行加入数据表中作为字段名
            for (int col = 1; col <= baseSheet.UsedRange.Columns.Count; col++)
            {
                baseTable.Columns.Add(baseSheet.Cells[1, col].value);

            }

            for (int row = 2; row <= baseSheet.UsedRange.Rows.Count; row++)
            {
                DataRow dr = baseTable.NewRow();
                for (int col = 1; col <= baseSheet.UsedRange.Columns.Count; col++)
                {
                    dr[col - 1] = baseSheet.Cells[row, col].value;

                }
              /*  if (dr["是否要求发电"].ToString() == "联通不要求发电")
                {
                    if (dr["运营商及共享"].ToString().Length > 2)
                    {
                        dr["共享数"] = Convert.ToInt16(dr["共享数"]) - 1;
                        dr["运营商及共享"] = dr["运营商及共享"].ToString().Replace("/联通", "");
                    }

                }*/
                baseTable.Rows.Add(dr);

            }

        }
        //生成发电记录表
        public void InitRecordTable()
        {
            Excel.Worksheet recordSheet = new Excel.Worksheet();
            recordSheet = Globals.ThisWorkbook.Worksheets["发电记录"];

            //将数据表第一行加入数据表中作为字段名
            for (int col = 1; col <= recordSheet.UsedRange.Columns.Count; col++)
            {
                recordTable.Columns.Add(recordSheet.Cells[1, col].value);

            }

            for (int row = 2; row <= recordSheet.UsedRange.Rows.Count; row++)
            {
                DataRow dr = recordTable.NewRow();
                for (int col = 1; col <= recordSheet.UsedRange.Columns.Count; col++)
                {
                    dr[col - 1] = recordSheet.Cells[row, col].value;

                }

                recordTable.Rows.Add(dr);

            }
        }
        #region VSTO 设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisWorkbook_Startup);
            this.Shutdown += new System.EventHandler(ThisWorkbook_Shutdown);
        }

        #endregion

    }
}
