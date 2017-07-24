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
    public partial class outputReport
    {

        private Excel.Worksheet baseSheet = new Excel.Worksheet();
        private DataTable baseTable = new DataTable();
        //发电记录表
        private Excel.Worksheet recordSheet = new Excel.Worksheet();
        private DataTable recordTable = new DataTable();
        //模板
        private Excel.Worksheet templetSheet = new Excel.Worksheet();

        private void outputReport_Startup(object sender, System.EventArgs e)
        {
            baseSheet = Globals.ThisWorkbook.Worksheets["基础信息"];
            baseTable = Globals.ThisWorkbook.baseTable;
            recordSheet = Globals.ThisWorkbook.Worksheets["发电记录"];
            recordTable = Globals.ThisWorkbook.recordTable;
            InitTempletCombox();

        }

        private void outputReport_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO 设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InternalStartup()
        {
            this.dailyReportBNT.Click += new System.EventHandler(this.dailyReportBNT_Click);
            this.reportStartTimePick.ValueChanged += new System.EventHandler(this.reportStartTimePick_ValueChanged);
            this.reportEndTimePicker.ValueChanged += new System.EventHandler(this.reportEndTimePicker_ValueChanged);
            this.reflashTempletBNT.Click += new System.EventHandler(this.reflashTempletBNT_Click);
            this.Startup += new System.EventHandler(this.outputReport_Startup);
            this.Shutdown += new System.EventHandler(this.outputReport_Shutdown);

        }


        #endregion








        //生成报表
        private void dailyReportBNT_Click(object sender, EventArgs e)
        {
            try
            {
                if (DateTime.Compare(reportStartTimePick.Value, reportEndTimePicker.Value) > 0)
                {
                    MessageBox.Show("结束时间大于起始时间");
                    return;

                }
                templetSheet = Globals.ThisWorkbook.Worksheets[templetComBox.Text];
                //根据时间和模板生成对应记录表
                DataTable recordTb = setRecordData(reportStartTimePick.Value, reportEndTimePicker.Value, templetComBox.Text);
                //根据模板生成报表
                switch (templetComBox.Text)
                {

                    case "日报表模板":
                        reportAccordTemplet(templetSheet, recordTb);

                        break;

                    case "电信周报模板":

                        reportAccordTemplet(templetSheet, recordTb);
                        setTitle();
                        if (setSumTable.Checked == true)
                        {
                            creatSumTable(recordTb);
                            this.Application.ActiveSheet.Cells[2, 6] = "其中电信分摊金额";
                        }
                        break;

                    case "移动周报表模板":
                        reportAccordTemplet(templetSheet, recordTb);
                        setTitle();
                        if (setSumTable.Checked == true)
                        {
                            creatSumTable(recordTb);
                            this.Application.ActiveSheet.Cells[2, 6] = "其中移动分摊金额";
                        }
                        break;
                    case "联通周报模板":
                        reportAccordTemplet(templetSheet, recordTb);
                        setTitle();
                        if (setSumTable.Checked == true) creatSumTable(recordTb);
                        break;
                    case "铁塔汇总表":
                        reportAccordTemplet(templetSheet, recordTb);
                        setTitle();
                        if (setSumTable.Checked == true) creatSumTable(recordTb);
                        break;
                    case "移动发电明细表":
                        reportAccordTemplet(templetSheet, recordTb);
                        setTitle();
                        if (setSumTable.Checked == true) creatSumTable(recordTb);
                        break;
                    default:
                        reportAccordTemplet(templetSheet, recordTb);
                        if (setSumTable.Checked == true) creatSumTable(recordTb);
                        break;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }
        //生成统计工作表
        private void creatSumTable(DataTable record)
        {
            this.Application.ScreenUpdating = false;
            Globals.ThisWorkbook.Worksheets["汇总表模板"].Copy(this.Application.ActiveSheet);

            Excel.Worksheet sumsheet = this.Application.ActiveSheet;
            sumsheet.Cells[1, 1] = " 崇仁县铁塔发电费结算周报汇总表（" +
                string.Format("{0:D}", reportStartTimePick.Value) + "至" +
               string.Format("{0:D}", reportEndTimePicker.Value) + ")";
            int toltalRowNum = 1000;


            switch (templetComBox.Text)
            {
                case "电信周报模板":

                    sumsheet.Cells[3, 2] = "=COUNTIF(电信周报模板!S4:S" + toltalRowNum + ",\"存量移动\")";
                    sumsheet.Cells[3, 3] = "=SUMIF(电信周报模板!S4:S" + toltalRowNum + ",\"存量移动\",电信周报模板!U4:U" + toltalRowNum + ")";
                    sumsheet.Cells[3, 6] = "=SUMIF(电信周报模板!S4:S" + toltalRowNum + ",\"存量移动\",电信周报模板!z4:z" + toltalRowNum + ")";
                    sumsheet.Cells[4, 2] = "=COUNTIF(电信周报模板!S4:S" + toltalRowNum + ",\"存量电信\")";
                    sumsheet.Cells[4, 3] = "=SUMIF(电信周报模板!S4:S" + toltalRowNum + ",\"存量电信\",电信周报模板!U4:U" + toltalRowNum + ")";
                    sumsheet.Cells[4, 6] = "=SUMIF(电信周报模板!S4:S" + toltalRowNum + ",\"存量电信\",电信周报模板!z4:z" + toltalRowNum + ")";
                    sumsheet.Cells[5, 2] = "=COUNTIF(电信周报模板!S4:S" + toltalRowNum + ",\"存量联通\")";
                    sumsheet.Cells[5, 3] = "=SUMIF(电信周报模板!S4:S" + toltalRowNum + ",\"存量联通\",电信周报模板!U4:U" + toltalRowNum + ")";
                    sumsheet.Cells[5, 6] = "=SUMIF(电信周报模板!S4:S" + toltalRowNum + ",\"存量联通\",电信周报模板!z4:z" + toltalRowNum + ")";
                    sumsheet.Cells[6, 2] = "=COUNTIF(电信周报模板!S4:S" + toltalRowNum + ",\"铁塔新建\")";
                    sumsheet.Cells[6, 3] = "=SUMIF(电信周报模板!S4:S" + toltalRowNum + ",\"铁塔新建\",电信周报模板!U4:U" + toltalRowNum + ")";
                    sumsheet.Cells[6, 6] = "=SUMIF(电信周报模板!S4:S" + toltalRowNum + ",\"铁塔新建\",电信周报模板!z4:z" + toltalRowNum + ")";
                    break;
                case "移动周报表模板":
                    sumsheet.Cells[3, 2] = "=COUNTIF(移动周报表模板!v4:v" + toltalRowNum + ",\"存量移动\")";
                    sumsheet.Cells[3, 3] = "=SUMIF(移动周报表模板!v4:v" + toltalRowNum + ",\"存量移动\",移动周报表模板!p4:p" + toltalRowNum + ")";
                    sumsheet.Cells[3, 6] = "=SUMIF(移动周报表模板!v4:v" + toltalRowNum + ",\"存量移动\",移动周报表模板!u4:u" + toltalRowNum + ")";
                    sumsheet.Cells[4, 2] = "=COUNTIF(移动周报表模板!v4:v" + toltalRowNum + ",\"存量电信\")";
                    sumsheet.Cells[4, 3] = "=SUMIF(移动周报表模板!v4:v" + toltalRowNum + ",\"存量电信\",移动周报表模板!p4:p" + toltalRowNum + ")";
                    sumsheet.Cells[4, 6] = "=SUMIF(移动周报表模板!v4:v" + toltalRowNum + ",\"存量电信\",移动周报表模板!u4:u" + toltalRowNum + ")";
                    sumsheet.Cells[5, 2] = "=COUNTIF(移动周报表模板!v4:v" + toltalRowNum + ",\"存量联通\")";
                    sumsheet.Cells[5, 3] = "=SUMIF(移动周报表模板!v4:v" + toltalRowNum + ",\"存量联通\",移动周报表模板!p4:p" + toltalRowNum + ")";
                    sumsheet.Cells[5, 6] = "=SUMIF(移动周报表模板!v4:v" + toltalRowNum + ",\"存量联通\",移动周报表模板!u4:u" + toltalRowNum + ")";
                    sumsheet.Cells[6, 2] = "=COUNTIF(移动周报表模板!v4:v" + toltalRowNum + ",\"铁塔新建\")";
                    sumsheet.Cells[6, 3] = "=SUMIF(移动周报表模板!v4:v" + toltalRowNum + ",\"铁塔新建\",移动周报表模板!p4:p" + toltalRowNum + ")";
                    sumsheet.Cells[6, 6] = "=SUMIF(移动周报表模板!v4:v" + toltalRowNum + ",\"铁塔新建\",移动周报表模板!u4:u" + toltalRowNum + ")";
                    break;
                case "移动发电明细表":
                    sumsheet.Cells[3, 2] = "=COUNTIF(移动发电明细表!ac3:ac" + toltalRowNum + ",\"存量移动\")";
                    sumsheet.Cells[3, 3] = "=SUMIF(移动发电明细表!ac3:ac" + toltalRowNum + ",\"存量移动\",移动发电明细表!w3:w" + toltalRowNum + ")";
                    sumsheet.Cells[3, 6] = "=SUMIF(移动发电明细表!ac3:ac" + toltalRowNum + ",\"存量移动\",移动发电明细表!ab3:ab" + toltalRowNum + ")";
                    sumsheet.Cells[4, 2] = "=COUNTIF(移动发电明细表!ac3:ac" + toltalRowNum + ",\"存量电信\")";
                    sumsheet.Cells[4, 3] = "=SUMIF(移动发电明细表!ac3:ac" + toltalRowNum + ",\"存量电信\",移动发电明细表!w3:w" + toltalRowNum + ")";
                    sumsheet.Cells[4, 6] = "=SUMIF(移动发电明细表!ac3:ac" + toltalRowNum + ",\"存量电信\",移动发电明细表!ab3:ab" + toltalRowNum + ")";
                    sumsheet.Cells[5, 2] = "=COUNTIF(移动发电明细表!ac3:ac" + toltalRowNum + ",\"存量联通\")";
                    sumsheet.Cells[5, 3] = "=SUMIF(移动发电明细表!ac3:ac" + toltalRowNum + ",\"存量联通\",移动发电明细表!w3:w" + toltalRowNum + ")";
                    sumsheet.Cells[5, 6] = "=SUMIF(移动发电明细表!ac3:ac" + toltalRowNum + ",\"存量联通\",移动发电明细表!ab3:ab" + toltalRowNum + ")";
                    sumsheet.Cells[6, 2] = "=COUNTIF(移动发电明细表!ac3:ac" + toltalRowNum + ",\"铁塔新建\")";
                    sumsheet.Cells[6, 3] = "=SUMIF(移动发电明细表!ac3:ac" + toltalRowNum + ",\"铁塔新建\",移动发电明细表!w3:w" + toltalRowNum + ")";
                    sumsheet.Cells[6, 6] = "=SUMIF(移动发电明细表!ac3:ac" + toltalRowNum + ",\"铁塔新建\",移动发电明细表!ab3:ab" + toltalRowNum + ")";

                    break;
                case "铁塔汇总表":
                    sumsheet.Cells[3, 2] = "=COUNTIF(铁塔汇总表!u4:u" + toltalRowNum + ",\"存量移动\")";
                    sumsheet.Cells[3, 3] = "=SUMIF(铁塔汇总表!u4:u" + toltalRowNum + ",\"存量移动\",铁塔汇总表!o4:o" + toltalRowNum + ")";
                    sumsheet.Cells[3, 6] = "=SUMIF(铁塔汇总表!u4:u" + toltalRowNum + ",\"存量移动\",铁塔汇总表!t4:t" + toltalRowNum + ")";
                    sumsheet.Cells[4, 2] = "=COUNTIF(铁塔汇总表!u4:u" + toltalRowNum + ",\"存量电信\")";
                    sumsheet.Cells[4, 3] = "=SUMIF(铁塔汇总表!u4:u" + toltalRowNum + ",\"存量电信\",铁塔汇总表!o4:o" + toltalRowNum + ")";
                    sumsheet.Cells[4, 6] = "=SUMIF(铁塔汇总表!u4:u" + toltalRowNum + ",\"存量电信\",铁塔汇总表!t4:t" + toltalRowNum + ")";
                    sumsheet.Cells[5, 2] = "=COUNTIF(铁塔汇总表!u4:u" + toltalRowNum + ",\"存量联通\")";
                    sumsheet.Cells[5, 3] = "=SUMIF(铁塔汇总表!u4:u" + toltalRowNum + ",\"存量联通\",铁塔汇总表!o4:o" + toltalRowNum + ")";
                    sumsheet.Cells[5, 6] = "=SUMIF(铁塔汇总表!u4:u" + toltalRowNum + ",\"存量联通\",铁塔汇总表!t4:t" + toltalRowNum + ")";
                    sumsheet.Cells[6, 2] = "=COUNTIF(铁塔汇总表!u4:u" + toltalRowNum + ",\"铁塔新建\")";
                    sumsheet.Cells[6, 3] = "=SUMIF(铁塔汇总表!u4:u" + toltalRowNum + ",\"铁塔新建\",铁塔汇总表!o4:o" + toltalRowNum + ")";
                    sumsheet.Cells[6, 6] = "=SUMIF(铁塔汇总表!u4:u" + toltalRowNum + ",\"铁塔新建\",铁塔汇总表!t4:t" + toltalRowNum + ")";

                    break;
                default:

                    break;
            }

            this.Application.ScreenUpdating = true;


        }
        //生成模板和时间对应记录
        private DataTable setRecordData(DateTime startDate, DateTime endDate, string templetName)
        {
            DataTable recordtb = new DataTable();
            recordtb = recordTable.Clone();
            //if (endDate < startDate) return;

            for (int row = 0; row < recordTable.Rows.Count; row++)
            {
                if (DateTime.Compare(Convert.ToDateTime(recordTable.Rows[row]["停电时间"]), startDate) > 0 &&
                      DateTime.Compare(Convert.ToDateTime(recordTable.Rows[row]["停电时间"]), endDate) < 0)
                {
                    bool copyFlag = false;
                    if (baseTable.Rows[FindBaseTableIndex(recordTable.Rows[row]["站址编码"].ToString())]["运营商及共享"].ToString().Contains("电信") &&
                        IsGetTelecomDatacheckBox.Checked == true)
                        copyFlag = true;
                    if (baseTable.Rows[FindBaseTableIndex(recordTable.Rows[row]["站址编码"].ToString())]["运营商及共享"].ToString().Contains("移动") &&
                        IsGetMobileDatacheckBox.Checked == true)
                        copyFlag = true;
                    if (baseTable.Rows[FindBaseTableIndex(recordTable.Rows[row]["站址编码"].ToString())]["运营商及共享"].ToString().Contains("联通") &&
                        IsGetUnicomDatacheckBox.Checked == true)
                        copyFlag = true;

                    if (baseTable.Rows[FindBaseTableIndex(recordTable.Rows[row]["站址编码"].ToString())]["是否要求发电"].ToString().Contains("不要求发电"))
                        if (IsGetNotRequireStationcheckBox.Checked == true)
                        {
                            copyFlag = true;
                        }
                        else
                        {
                            copyFlag = false;
                        }


                    if (copyFlag)
                        recordtb.Rows.Add(copyRecordRow(recordTable, recordtb, row));



                }

            }



            /* switch (templetName)
             {

                 case "日报表模板":
                     for (int row = 0; row < recordTable.Rows.Count; row++)
                     {
                         if (DateTime.Compare(Convert.ToDateTime(recordTable.Rows[row]["停电时间"]), startDate) > 0 &&
                             DateTime.Compare(Convert.ToDateTime(recordTable.Rows[row]["停电时间"]), endDate) < 0)
                         {
                             recordtb.Rows.Add(copyRecordRow(recordTable, recordtb, row));
                         }
                     }
                     break;

                 case "电信周报模板":
                     for (int row = 0; row < recordTable.Rows.Count; row++)
                     {
                         if (DateTime.Compare(Convert.ToDateTime(recordTable.Rows[row]["停电时间"]), startDate) > 0 &&
                             DateTime.Compare(Convert.ToDateTime(recordTable.Rows[row]["停电时间"]), endDate) < 0)
                         {
                             if (baseTable.Rows[FindBaseTableIndex(recordTable.Rows[row]["站址编码"].ToString())]["运营商及共享"].ToString().Contains("电信"))
                                 recordtb.Rows.Add(copyRecordRow(recordTable, recordtb, row));
                         }
                     }
                     break;

                 case "移动周报表模板":
                     for (int row = 0; row < recordTable.Rows.Count; row++)
                     {
                         if (DateTime.Compare(Convert.ToDateTime(recordTable.Rows[row]["停电时间"]), startDate) > 0 &&
                             DateTime.Compare(Convert.ToDateTime(recordTable.Rows[row]["停电时间"]), endDate) < 0)
                         {
                             if (baseTable.Rows[FindBaseTableIndex(recordTable.Rows[row]["站址编码"].ToString())]["运营商及共享"].ToString().Contains("移动"))
                                 recordtb.Rows.Add(copyRecordRow(recordTable, recordtb, row));
                         }
                     }
                     break;
                 case "联通周报模板":
                     for (int row = 0; row < recordTable.Rows.Count; row++)
                     {
                         if (DateTime.Compare(Convert.ToDateTime(recordTable.Rows[row]["停电时间"]), startDate) > 0 &&
                             DateTime.Compare(Convert.ToDateTime(recordTable.Rows[row]["停电时间"]), endDate) < 0)
                         {
                             if (baseTable.Rows[FindBaseTableIndex(recordTable.Rows[row]["站址编码"].ToString())]["基站来源"].ToString().Contains("存量联通") &&
                                 baseTable.Rows[FindBaseTableIndex(recordTable.Rows[row]["站址编码"].ToString())]["是否要求发电"].ToString() != "联通不要求发电")
                                 recordtb.Rows.Add(copyRecordRow(recordTable, recordtb, row));
                         }
                     }
                     break;
                 default:
                     for (int row = 0; row < recordTable.Rows.Count; row++)
                     {
                         if (DateTime.Compare(Convert.ToDateTime(recordTable.Rows[row]["停电时间"]), startDate) > 0 &&
                             DateTime.Compare(Convert.ToDateTime(recordTable.Rows[row]["停电时间"]), endDate) < 0)
                         {
                             recordtb.Rows.Add(copyRecordRow(recordTable, recordtb, row));
                         }
                     }
                     break;
             }*/
            return recordtb;


        }
        //按模板生成报表
        private void reportAccordTemplet(Excel.Worksheet templet, DataTable recordTb)
        {
            this.Application.ScreenUpdating = false;
            //复制模板到新的工作簿


            Excel.Workbook newWorkbook = this.Application.Workbooks.Add();
            templet.Copy(newWorkbook.Worksheets[1]);
            //newWorkbook.Application.Caption = templetComBox.Text;
            Excel.Worksheet newWorkSheet = newWorkbook.Worksheets[templetComBox.Text];
            //删除自动生成的SHEET1
            newWorkbook.Worksheets["Sheet1"].Delete();
            newWorkSheet.Activate();


            for (int row = 0; row < recordTb.Rows.Count; row++)
            {
                int startRow = newWorkSheet.UsedRange.Rows.Count + 1;
                //在基础数据表中查找对应站址编码的位置
                int baseTableIndex = FindBaseTableIndex(recordTb.Rows[row]["站址编码"].ToString());

                for (int col = 1; col <= templet.UsedRange.Columns.Count; col++)
                {

                    string tempString = templet.Cells[1, col].value.ToString();
                    switch (tempString)
                    {
                        case "常量":
                            newWorkSheet.Cells[startRow, col] = templet.Cells[2, col];
                            break;
                        case "计算":
                            newWorkSheet.Cells[2, col].copy();
                            newWorkSheet.Cells[startRow, col].PasteSpecial();

                            break;
                        case "发电记录":
                            newWorkSheet.Cells[startRow, col] = recordTb.Rows[row][templet.Cells[2, col].value];
                            break;
                        case "基础信息":
                            newWorkSheet.Cells[startRow, col] = baseTable.Rows[baseTableIndex][templet.Cells[2, col].value.ToString()];
                            break;
                        default:
                            break;

                    }
                }

            }
            //删除模板上面两行
            newWorkSheet.Rows[1].Delete();
            newWorkSheet.Rows[1].Delete();

            //if (templetComBox.Text != "日报表模板")




            this.Application.ScreenUpdating = true;
        }
        //生成表头
        private void setTitle()
        {
            this.Application.ActiveSheet.Cells[1, 1] = " 崇仁县铁塔发电费结算周报明细表（" +
                                                        string.Format("{0:D}", reportStartTimePick.Value) + "至" +
                                                        string.Format("{0:D}", reportEndTimePicker.Value) + ")";
        }




        //辅助函数》》》》

        //复制数据行
        private DataRow copyRecordRow(DataTable sourse, DataTable target, int row)
        {
            DataRow dr = target.NewRow();
            for (int i = 0; i < sourse.Columns.Count; i++)
            {
                dr[i] = sourse.Rows[row][i];
            }
            return dr;

        }
        //在基础数据表中查找对应站址编码的位置
        private int FindBaseTableIndex(string stationID)
        {
            int baseTableIndex = -1;

            for (int i = 0; i < baseTable.Rows.Count; i++)
            {
                if (baseTable.Rows[i]["站址编码"].ToString() == stationID)
                {
                    baseTableIndex = i;
                    return baseTableIndex;
                }


            }
            if (baseTableIndex == -1)
            {
                MessageBox.Show("未找到该站点,请更新基础数据表");

            }
            return baseTableIndex;
        }
        //根据字段名查找列数
        private int FindColumnIndex(Excel.Worksheet sheet, int row, string columnName)
        {
            int result = -1;
            for (int i = 1; i <= sheet.UsedRange.Columns.Count; i++)
            {
                if (sheet.Cells[row, i] == columnName)
                {
                    result = i;
                }
            }
            return result;
        }

        //辅助函数》》》》



        //初始化模板选择框
        private void InitTempletCombox()
        {
            this.templetComBox.Items.Clear();
            string[] exceptName = { "输入数据", "生成报表", "发电记录", "基础信息", "汇总表模板" };
            foreach (Excel.Worksheet ws in Globals.ThisWorkbook.Worksheets)
            {
                Random a = new Random();
                if (ws.Name == "Sheet1") ws.Name = "Sheet" + a.Next();
                bool flag = true;
                for (int i = 0; i < exceptName.Length; i++)
                {
                    if (ws.Name == exceptName[i])
                    {
                        flag = false;
                    }

                }
                if (flag)
                {
                    this.templetComBox.Items.Add(ws.Name);
                }
            }
        }

















        private void reportStartTimePick_ValueChanged(object sender, EventArgs e)
        {
            string tempDt;
            tempDt = reportStartTimePick.Value.ToShortDateString().ToString();
            tempDt += " 00:00:00";
            reportStartTimePick.Value = Convert.ToDateTime(tempDt);
        }

        private void reportEndTimePicker_ValueChanged(object sender, EventArgs e)
        {
            string tempDt;
            tempDt = reportEndTimePicker.Value.ToShortDateString().ToString();
            tempDt += " 23:59:59";
            reportEndTimePicker.Value = Convert.ToDateTime(tempDt);

        }

        private void reflashTempletBNT_Click(object sender, EventArgs e)
        {
            InitTempletCombox();
        }
    }
}

