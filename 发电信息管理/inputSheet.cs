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
using Microsoft.Office.Interop.Excel;
using System.Collections;
using System.IO;

namespace 发电信息管理
{
    public partial class inputSheet
    {
        //基础数据记录表
        private System.Data.DataTable baseTable = new System.Data.DataTable("baseTable");
        private System.Data.DataTable temptable = new System.Data.DataTable("tempTable");
        private Excel.Worksheet baseSheet = new Excel.Worksheet();
        //发电记录表
        private Excel.Worksheet recordSheet = new Excel.Worksheet();
        private System.Data.DataTable recordTable = new System.Data.DataTable("recordTable");



        private void Sheet1_Startup(object sender, System.EventArgs e)
        {
            baseSheet = Globals.ThisWorkbook.Worksheets["基础信息"];
            baseTable = Globals.ThisWorkbook.baseTable;
            recordSheet = Globals.ThisWorkbook.Worksheets["发电记录"];
            recordTable = Globals.ThisWorkbook.recordTable;
            InitComboBox(this.stationComBox);
        }

        private void Sheet1_Shutdown(object sender, System.EventArgs e)
        {
        }



        private void InitComboBox(ComboBox combobox)
        {
            try
            {
                combobox.DrawMode = DrawMode.Normal;//设置绘制方式
                combobox.DropDownHeight = 60;
                combobox.DropDownStyle = ComboBoxStyle.DropDown;
                combobox.FlatStyle = FlatStyle.Flat; //设置外观
                                                     /*一般应用
                                                      combobox.Items.Add("hello");
                                                      combobox.Items.Add("word");
                                                    */
                                                     //数据绑定方式


                combobox.DataSource = baseTable;//设置数据源
                combobox.DisplayMember = "基站站名";//设置显示列
                combobox.ValueMember = "站址编码";//设置实际值
                combobox.SelectedIndexChanged += new EventHandler(comboBox1_SelectedIndexChanged);
                this.textBox2.Text = this.stationComBox.SelectedValue.ToString();

            }
            catch (Exception ex)
            {

            }
        }
        #region VSTO 设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InternalStartup()
        {
            this.stationComBox.SelectedIndexChanged += new System.EventHandler(this.comboBox1_SelectedIndexChanged);
            this.losePowerTimePicker.ValueChanged += new System.EventHandler(this.losePowerTimePicker_ValueChanged);
            this.textBox1.TextChanged += new System.EventHandler(this.textBox1_TextChanged);
            this.submitBnt.Click += new System.EventHandler(this.submitBnt_Click);
            this.selectImageBNT.Click += new System.EventHandler(this.selectImageBNT_Click);
            this.Startup += new System.EventHandler(this.Sheet1_Startup);
            this.Shutdown += new System.EventHandler(this.Sheet1_Shutdown);

        }

        #endregion

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            this.textBox2.Text = this.stationComBox.SelectedValue.ToString();
        }



        //文本框输入改变事件

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            temptable = new System.Data.DataTable();
            temptable.Columns.Add("基站站名", typeof(string));
            temptable.Columns.Add("站址编码", typeof(string));
            if (string.IsNullOrEmpty(this.textBox1.Text))
            {
                //当Combobox输入为空时，将列表所有项加载进来
                temptable = baseTable;
                return;
            }
            else
            {

                for (int row = 0; row <= baseTable.Rows.Count - 1; row++)
                {

                    if (baseTable.Rows[row]["基站站名"].ToString().Contains(this.textBox1.Text))
                    {
                        DataRow dr = temptable.NewRow();
                        dr["基站站名"] = baseTable.Rows[row]["基站站名"].ToString()+ "-"+baseTable.Rows[row]["基站来源"].ToString();
                        dr["站址编码"] = baseTable.Rows[row]["站址编码"].ToString();
                        temptable.Rows.Add(dr);
                    }
                }

            }
            this.stationComBox.DataSource = temptable;

            /*if (this.stationComBox.DroppedDown == false)
            {

                this.stationComBox.DroppedDown = true;

            }*/
            
        }
        //提交按钮事件

        private void submitBnt_Click(object sender, EventArgs e)

        {
            int usedRowNum = recordSheet.UsedRange.Rows.Count;
            DataRow dr = recordTable.NewRow();
            //id
            recordSheet.Cells[usedRowNum + 1, 1] = recordSheet.UsedRange.Rows.Count;
            dr["ID"] = recordSheet.Cells[usedRowNum + 1, 1];
            //县区
            recordSheet.Cells[usedRowNum + 1, 2] = "崇仁县";
            dr["县区"] = "崇仁县";
            //站点名称
            recordSheet.Cells[usedRowNum + 1, 3] = stationComBox.Text.ToString();
            dr["站点名称"] = stationComBox.Text.ToString();
            //设置单元格式为文本，站址编码
            recordSheet.Cells[usedRowNum + 1, 4].NumberFormatLocal = "@";
            recordSheet.Cells[usedRowNum + 1, 4] = stationComBox.SelectedValue.ToString();
            dr["站址编码"] = stationComBox.SelectedValue.ToString();
            //停电时间
            recordSheet.Cells[usedRowNum + 1, 5] = losePowerTimePicker.Value;
            dr["停电时间"] = losePowerTimePicker.Value;
            //发电开始时间
            recordSheet.Cells[usedRowNum + 1, 6] = startGenerationTimePicker.Value;
            dr["发电起始时间"] = startGenerationTimePicker.Value;
            //发电终止时间
            recordSheet.Cells[usedRowNum + 1, 7] = stopGenerationTimePicker.Value;
            dr["发电终止时间"] = stopGenerationTimePicker.Value;
            //发电前电压
            recordSheet.Cells[usedRowNum + 1, 8] = beforeGenVoltage.Text;
            dr["发电前电压"] = beforeGenVoltage.Text;

            recordTable.Rows.Add(dr);
        }

        private void losePowerTimePicker_ValueChanged(object sender, EventArgs e)
        {
            startGenerationTimePicker.Value = losePowerTimePicker.Value;
            stopGenerationTimePicker.Value = losePowerTimePicker.Value;
        }

        private void openFileDialog_FileOk(object sender, System.ComponentModel.CancelEventArgs e)
        {

        }



        private void selectImageBNT_Click(object sender, EventArgs e)
        {
            /*OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.InitialDirectory = "c:\\";//注意这里写路径时要用c:\\而不是c:\
            openFileDialog.Filter = "图片文件|*.jpg|所有文件|*.*|bmp文件|*.bmp";
            openFileDialog.RestoreDirectory = true;
            openFileDialog.FilterIndex = 1;
            openFileDialog.Multiselect = true;
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                string[] filenames = openFileDialog.FileNames;
                foreach ( string filename in filenames)
                {
                    this.Shapes.AddPicture(filename, Office.MsoTriState.msoFalse, Office.MsoTriState.msoTrue, 100, 100, 100, 100);
                }
                 
            }*/

            FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog();
            if(folderBrowserDialog.ShowDialog() == DialogResult.OK)
            {
                string dirname = folderBrowserDialog.SelectedPath;
                string[] fileType = { ".jpg", ".jpeg", ".bmp", ".png", ".gif" };
                setPicture(dirname, fileType);
            }
        }
        public void setPicture(string path, string[] exception)
        {
            int row = 1;

            Excel.Worksheet pictureSt = this.Application.Workbooks.Add().Worksheets[1];
            pictureSt.Name = "图片";

            DirectoryInfo search = new DirectoryInfo(path);   
            //获取目录path下所有目录和子目录下的文件
            FileSystemInfo[] fsinfos = search.GetFileSystemInfos();
            //遍历目录下文件及其子目录下文件
            foreach (FileSystemInfo fsinfo in fsinfos)
            {
                
                if (fsinfo is DirectoryInfo)     //判断是否为文件夹
                {

                    pictureSt.Cells[row, 1]= fsinfo.Name;
                    row++;
                    for(int i = 1; i < 6; i++)
                    {

                    }
                    DirectoryInfo dtinfo = new DirectoryInfo(fsinfo.FullName);
                    FileInfo[] f = dtinfo.GetFiles();//获取子目录下的文件
                    int col = 1;
                    foreach (FileInfo file in f)
                    {
                        
                        
                        for (int i = 0; i < exception.Length; i++)
                        {
                            if (file.Name.Contains(exception[i]) == true)
                            {
                                pictureSt.Cells[row, col] = file.Name;
                                pictureSt.Cells[row+1, col].ColumnWidth = 20;
                                pictureSt.Cells[row+1,col].RowHeight = 100;

                                pictureSt.Shapes.AddPicture(file.FullName, Office.MsoTriState.msoFalse, Office.MsoTriState.msoTrue,
                                    pictureSt.Cells[row+1, col].Left, pictureSt.Cells[row+1, col].Top, 100,100);
                                    col++;
                            }
                        }

                    }
                    row+=2;
                }
                else
                {
                    for (int i = 0; i < exception.Length; i++)
                    {
                        if (fsinfo.Name.Contains(exception[i]) == true)
                        {

                            //Item.Add(fsinfo.FullName);
                            //j++;
                        }
                    }

                }
            }

            //return ;
        }
    }
}
