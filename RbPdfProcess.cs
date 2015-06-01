using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using System.Diagnostics;
using System.IO;
using ExcelAddIn4Pdf.Util;
using System.Text.RegularExpressions;
using System.Data;
using System.Data.OleDb;
using pdf;
using System.Collections;

namespace ExcelAddIn4Pdf
{
    public partial class RbPdfProcess
    {       
        int rowfix = 0;
        private string configPath = System.AppDomain.CurrentDomain.BaseDirectory + "Config\\PTC.xls";

        public DataSet dsConfig { get; set; } 

        public string strPdfInfo = "";//文件中读出并处理过的字符
        public int numOfPeople = 0;//人数
        public int Vehicle = 0;//车座
        public int iLineStart = 16;
        public int iLine = 16;//行数

        public Dictionary<string,string> dicBookInfo = new Dictionary<string,string>(); //订单信息
        public SortedList<DateTime, string> dicDays = new SortedList<DateTime, string>();//日期信息
        public SortedList<DateTime, SortedList<DateTime, string>> dicDayTravel = new SortedList<DateTime, SortedList<DateTime, string>>();//日期时间信息

        private List<CairnsTravel> dicDayInCairns = new List<CairnsTravel>();//日期时间信息
        
        private MDSettings mdConfig = MDSettings.getInstance();

        #region 私有方法
        /// <summary>加载Excel</summary>
        /// <param name="filePath">Excel路径</param>
        /// <returns>Excel数据</returns>
        private DataSet LoadDataFromExcel(string filePath)
        {
            try
            {
                string strConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filePath + ";Extended Properties='Excel 8.0;HDR=False;IMEX=1'";
                OleDbConnection OleConn = new OleDbConnection(strConn);
                OleConn.Open();

                System.Data.DataTable dtName = OleConn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "Table" });
                DataSet OleDsExcle = new DataSet();
                foreach (DataRow dr in dtName.Rows)
                {
                    string sheetName = dr["TABLE_NAME"].ToString();
                    String sql = "SELECT * FROM [" + sheetName + "]";
                    OleDbDataAdapter OleDaExcel = new OleDbDataAdapter(sql, OleConn);
                    OleDaExcel.Fill(OleDsExcle, sheetName);
                }
                OleConn.Close();
                return OleDsExcle;
            }

            catch (Exception err)
            {
                //MessageBox.Show("数据绑定Excel失败!失败原因：" + err.Message, "提示信息",
                //    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return null;
            }
        }

        /// <summary> 获取pdf文档的信息</summary>
        /// <param name="strPdf"></param>
        private void AnalysisInformation(string strPdf)
        {
            StringAnalysis sd = new StringAnalysis();
            string[] sDays = null;
            if (strPdf.Contains("*")) sDays = strPdf.Split('*');
            if (sDays == null) return;
            foreach (var sDay in sDays)
            {
                if (sDay.Contains("#"))//截取日期信息
                {
                    DateTime dt = sd.GetDays(sDay);
                    dicDays.Add(dt, sDay);

                    SortedList<DateTime, string> dtTravel = sd.GetTravel(dt, sDay);
                    dicDayTravel.Add(dt, dtTravel);

                    if (sDay.Contains("&"))
                    {
                        int sWhere = sDay.IndexOf("&");
                        int eWhere = sDay.IndexOf("%");
                        if (eWhere == -1)//不存在%VehicleS时
                            eWhere = sDay.IndexOf("#");
                        if(Vehicle==0)
                        {
                            string strVehicle = sDay.Substring(eWhere+1,2);
                            int.TryParse(strVehicle.TrimEnd('S'), out Vehicle);
                        }

                        string strWhere = sDay.Substring(0, eWhere).Replace("&", ""); ;//获取地点
                        
                        if (strWhere.ToUpper().Contains("CAIRNS"))
                        {
                            dicDayInCairns.Add(new CairnsTravel( strWhere, dtTravel));
                        }
                    }
                }
                else
                {
                    string sInfo = sDay.Replace("  ", "@").Replace("\r\n", "");
                    string[] sBkInfo = sInfo.Split('@');
                    foreach (var line in sBkInfo)
                    {
                        if (!line.Contains(":")) continue;
                        int i = line.IndexOf(':');
                        string strKey = line.Substring(0, i).Trim();
                        string strValue = line.Substring(i + 1).Trim();
                        dicBookInfo.Add(strKey, strValue);

                        if (strKey == "Pax No")
                        {
                            string strNum = Regex.Replace(strValue, @"[^\d.\d]", " ");//剔除除数字外字符替换为空格
                            strNum = Regex.Replace(strNum, @"\s+", " ");//多个空格换位一个
                            string[] sNum = strNum.Split(' ');
                            foreach (var num in sNum)
                            {
                                int iNum = 0;
                                int.TryParse(num, out iNum);
                                numOfPeople += iNum;
                            }
                        }
                    }
                }
            }
        }

        /// <summary>静态信息</summary>
        /// <param name="workSheet"></param>
        private void WriteRegularInfo(Worksheet workSheet)
        {
            Range range = null;

            TitleStyle(workSheet, "G1", "I1", "Quotation",false,15);   
            TitleStyle(workSheet, "B1", "F1", "Tropcial Wonderland Tour Services",false,15); 

            #region Title:Picture
            string picPath = System.AppDomain.CurrentDomain.BaseDirectory + "Image\\logo.png";
            range = workSheet.get_Range("A1", "A5"); //获取Excel多个单元格区域：本例做为Excel表头    
            range.Merge(0); //单元格合并动作     
            range.ColumnWidth = 13;
            Microsoft.Office.Interop.Excel.Pictures pics = workSheet.Pictures();
            pics.Insert(picPath, range);  
            #endregion

            #region 公司信息1
            range = workSheet.get_Range("B2", "I5"); //获取Excel多个单元格区域：本例做为Excel表头   
            //range.Merge(0); //单元格合并动作  

            range.Cells[1, 1] = mdConfig.DicConfig["PTC"]["REG"];//Excel单元格赋值
            range.Cells[1, 5] = mdConfig.DicConfig["PTC"]["ABN"];
            range.Cells[1, 3] = mdConfig.DicConfig["PTC"]["ADDR"];
            range.Cells[2, 1] = mdConfig.DicConfig["PTC"]["MOBILE"];
            range.Cells[2, 4] = mdConfig.DicConfig["PTC"]["TEL"];
            range.Cells[3, 1] = mdConfig.DicConfig["PTC"]["EMAIL"];
            range.Cells[3, 1] = mdConfig.DicConfig["PTC"]["EMAIL"];
            range.Cells[4, 1] = mdConfig.DicConfig["PTC"]["FAX"];
            
            range.Font.Size = 10; //设置字体大小   
            range.Font.Underline = false; //设置字体是否有下划线   
            range.HorizontalAlignment = XlHAlign.xlHAlignLeft; //设置字体在单元格内的对其方式

            //range.Borders.LineStyle = 1; //设置单元格边框的粗细  
            range.BorderAround2(1);
            #endregion

            TitleStyle(workSheet, "B6", "I6", mdConfig.DicConfig["PTC"]["QLD"],false,15); 

            #region 客户信息1
            range = workSheet.get_Range("B7", "I7"); //获取Excel多个单元格区域：本例做为Excel表头   
            //range.Merge(0); //单元格合并动作  

            range.Cells[1, 1] = "No.of people";//Excel单元格赋值

            range.Cells[1, 3] = "Arrival Date";//Excel单元格赋值

            range.Cells[1, 5] = "Job  Number";//Excel单元格赋值

            range.Font.Size = 11; //设置字体大小   
            range.Font.Underline = false; //设置字体是否有下划线   QLD Transport Operator Approve Number:900273711
            range.HorizontalAlignment = XlHAlign.xlHAlignLeft; //设置字体在单元格内的对其方式

            //【订单信息】
            range = workSheet.get_Range("B7", "C8"); //获取Excel多个单元格区域：本例做为Excel表头   
            range.BorderAround2(1);

            range = workSheet.get_Range("D7", "E8"); //获取Excel多个单元格区域：本例做为Excel表头   
            range.BorderAround2(1);

            range = workSheet.get_Range("F7", "G8"); //获取Excel多个单元格区域：本例做为Excel表头   
            range.BorderAround2(1);

            range = workSheet.get_Range("B8", "I8"); //获取Excel多个单元格区域：本例做为Excel表头   
            range.NumberFormatLocal = "@";

            range.Cells[1, 1] = numOfPeople;//Excel单元格赋值

            range.Cells[1, 3] = dicDayInCairns[0].DayTravel.Keys[0].Date.ToString("yyyy/MM/dd");//Excel单元格赋值

            range.Cells[1, 5] = mdConfig.DicConfig["PTC"]["NO"].PadLeft(7, '0');//Excel单元格赋值

            int intNo=0;
            int.TryParse(mdConfig.DicConfig["PTC"]["NO"],out intNo);
            mdConfig.SetKeyValue("PTC","NO",(intNo+1).ToString());

            range.Font.Size = 10; //设置字体大小   
            range.Font.Underline = false; //设置字体是否有下划线   QLD Transport Operator Approve Number:900273711
            range.HorizontalAlignment = XlHAlign.xlHAlignLeft; //设置字体在单元格内的对其方式
            #endregion

            #region 客户信息1
            TitleStyle(workSheet, "A9", "E9", "Customer Information       ABN:", true,12);

            range = workSheet.get_Range("A10", "E14");
            range.Cells[1, 1] = mdConfig.DicConfig["PTC"]["PTCNAME"];//Excel单元格赋值

            range.Cells[2, 1] = mdConfig.DicConfig["PTC"]["PTCADDR1"];//Excel单元格赋值 PTCAddr
            range.Cells[3, 1] = mdConfig.DicConfig["PTC"]["PTCADDR2"];

            range.Cells[4, 1] = mdConfig.DicConfig["PTC"]["PTCPHONE"];//Excel单元格赋值 PTCPhone
            range.Cells[4, 4] = mdConfig.DicConfig["PTC"]["PTCFAX"];//Excel单元格赋值PTCFax

            range.Font.Size = 10; //设置字体大小   
            range.Font.Underline = false; //设置字体是否有下划线   QLD Transport Operator Approve Number:900273711
            range.HorizontalAlignment = XlHAlign.xlHAlignLeft; //设置字体在单元格内的对其方式
            range.BorderAround2(1);

            TitleStyle(workSheet, "G10", "G10", "Agent Ref:", true,10);
            TitleStyle(workSheet, "G12", "G12", "Require:", true,10);
            TitleStyle(workSheet, "G14", "G14", "Vehicle:", true,10);

            TitleStyle(workSheet, "H10", "H10", dicBookInfo["Booking Ref"], true, 10);
            TitleStyle(workSheet, "H12", "H12", "", true, 10);
            TitleStyle(workSheet, "H14", "H14", Vehicle.ToString(), true, 10);

            #endregion

            #region 旅游信息

            TitleStyle(workSheet, "A15", "C15", "Itinerary", true,15);

            TitleStyle(workSheet, "D15", "E15", "Quote Items", true, 15);

            TitleStyle(workSheet, "F15", "F15", "U-Price", true, 15);
            TitleStyle(workSheet, "G15", "G15", "Qty", true, 15);
            TitleStyle(workSheet, "H15", "H15", "Discount", true, 12);
            TitleStyle(workSheet, "I15", "I15", "T-Price", true, 15);

            #endregion
        }
        /// <summary>
        /// 写入凯恩斯内容
        /// </summary>
        /// <param name="workSheet"></param>
        private void WriteCairnsTravelInfo(Worksheet workSheet)
        {
            bool inCairns = false;
            int endCairnsLine = 0;
            for (int  k = 0;k< dicDayInCairns.Count;k++)
            {
                string rX = "A" + iLine++;
                string rY = "C" + iLine++;
                Range range = workSheet.get_Range(rX, rY);
                range.Merge(0);
                range.Font.Size = 11;
                range.HorizontalAlignment = XlHAlign.xlHAlignLeft; //设置字体在单元格内的对其方式
                range.Font.Name = "黑体"; //设置字体的种类  
                range.Cells[1][1].WrapText = true;
                range.EntireRow.AutoFit();
                range.Cells[1][1].Value = dicDayInCairns[k].DayTitle.Replace("\r","").Replace("\n","");
                 
                SortedList<DateTime, string> ICainsTravel = dicDayInCairns[k].DayTravel;

                GetQuoteAndPrice(ICainsTravel,iLine);
                for (int j = 0; j < ICainsTravel.Count; j++)
                {
                    if (k == 0 && (ICainsTravel.Values[j].ToUpper().Contains("ARRIVE CAIRNS"))) inCairns = true;//凯恩斯开始
                    if (k == dicDayInCairns.Count-1) //凯恩斯结束
                    {
                        for (int i=0 ;i< ICainsTravel.Count;i++)
                        {
                            if (ICainsTravel.Values[i].ToUpper().Contains("ARRIVE CAIRNS"))
                                endCairnsLine = i;
                        }
                    }

                    if (inCairns && k < dicDayInCairns.Count - 1)//前几天
                    {
                        rX = "A" + iLine++;
                        rY = "C" + iLine++;
                        range = workSheet.get_Range(rX, rY);
                        range.Merge(0);
                        range.Font.Size = 9;
                        range.EntireRow.AutoFit();
                        range.Cells[1][1].WrapText = true;
                        range.HorizontalAlignment = XlHAlign.xlHAlignLeft;
                        string sContent = Regex.Replace(ICainsTravel.Values[j].Trim(), @"[\u4e00-\u9fa5]", ""); //除去中文
                        sContent = Regex.Replace(sContent, @"\([^\(]*\)", "").Replace("\r","").Replace("\n","");
                        range.Cells[1][1].Value = ICainsTravel.Keys[j].ToString("HH:mm ") + sContent;
                    }
                    else if (inCairns && k == dicDayInCairns.Count - 1 && j <= endCairnsLine)//最后一天
                    {
                            rX = "A" + iLine++;
                            rY = "C" + iLine++;
                            range = workSheet.get_Range(rX, rY);
                            range.Merge(0);
                            range.Font.Size = 9;
                            range.EntireRow.AutoFit();
                            range.Cells[1][1].WrapText = true;
                            range.HorizontalAlignment = XlHAlign.xlHAlignLeft;
                            string sContent = Regex.Replace(ICainsTravel.Values[j].Trim(), @"[\u4e00-\u9fa5]", ""); //除去中文
                            sContent = Regex.Replace(sContent, @"\([^\(]*\)", "").Replace("\r", "").Replace("\n", "");
                            range.Cells[1][1].Value = ICainsTravel.Keys[j].ToString("HH:mm ") + sContent;
                            //range.Cells[1][1].Value = ICainsTravel.Keys[j].ToString("HH:mm ") + ICainsTravel.Values[j].Trim();
                    }
                }
                rX = "A" + iLine;
                rY = "C" + iLine++;
                range = workSheet.get_Range(rX, rY);
                range.Merge(0);
            }

            for (int i = iLineStart; i < iLine; i++)
            {
                workSheet.Cells[i, 9] = "=IF(F" + i + "*G" + i + "=0,IF(F" + i + "*(1-H" + i + ")=0,\"\",F" + i + "*(1-H" + i + ")),F" + i + "*G" + i + "*(1-H" + i + "))"; //PRODUCT(F" + i + ",G" +i+ ")
                workSheet.Cells[i, 9].NumberFormatLocal = "0.00";
            }

            TitleStyle(workSheet, "G" + iLine, "H" + iLine, "T-Price GST Inc:", true, 12);
            TitleStyle(workSheet, "I" + iLine, "I" + iLine, "=SUM(I" + iLineStart + ":I" + (iLine-1) + ")", true, 11);

            NoteStyle(workSheet, "A" + (iLine + 2), "E" + (iLine + 6), mdConfig.DicConfig["PTC"]["NOTE1"], false, 10);
            NoteStyle(workSheet, "A" + (iLine + 8), "E" + (iLine + 10), mdConfig.DicConfig["PTC"]["NOTE2"], false, 10);
            NoteStyle(workSheet, "A" + (iLine + 12), "E" + (iLine + 12), mdConfig.DicConfig["PTC"]["NOTE3"], false, 10);
            NoteStyle(workSheet, "A" + (iLine + 13), "E" + (iLine + 13), mdConfig.DicConfig["PTC"]["NOTE4"], false, 10);

            NoteStyle(workSheet, "G" + (iLine + 7), "I" + (iLine + 7), mdConfig.DicConfig["PTC"]["NOTE5"], false, 10);
            NoteStyle(workSheet, "G" + (iLine + 8), "I" + (iLine + 9), mdConfig.DicConfig["PTC"]["NOTE6"], false, 10);
            NoteStyle(workSheet, "G" + (iLine + 10), "I" + (iLine + 10), mdConfig.DicConfig["PTC"]["NOTE7"], false, 10);
            
        }   

        /// <summary>
        /// 获取收费项
        /// </summary>
        /// <param name="ICainsTravel"></param>
        /// <param name="iLine"></param>
        private void GetQuoteAndPrice(SortedList<DateTime, string> ICainsTravel, int iLine)
        {
            Dictionary<string, int> dicprice = GetPriceList();
            foreach (var sContent in ICainsTravel.Values)
            {
                //if(sContent.Contains("Palm Cove"))

            }
        }

        /// <summary>初始化</summary>
        private void init()
        {
            #region 初始化
            strPdfInfo = "";//文件中读出并处理过的字符串
            dicDays = new SortedList<DateTime, string>();
            numOfPeople = 0;
            Vehicle = 0;//车座
            iLine = 16;

            dicBookInfo = new Dictionary<string, string>(); //订单信息
            dicDays = new SortedList<DateTime, string>();//日期信息
            dicDayTravel = new SortedList<DateTime, SortedList<DateTime, string>>();//日期时间信息

            dicDayInCairns = new List<CairnsTravel>();//日期时间信息
            #endregion
        }

        #endregion

        private string GetStringFromFile(string filePath)
        {
            PdfReader pdf = new PdfReader();
            return strPdfInfo =pdf.getString(filePath);
        }

        private void RbPdfProcess_Load(object sender, RibbonUIEventArgs e)
        {
            if (File.Exists(configPath))
                dsConfig = LoadDataFromExcel(configPath);
        }

        private void btnOpen_Click(object sender, RibbonControlEventArgs e)
        {
            using (OpenFileDialog ofd = new OpenFileDialog())
            {
                ofd.Filter = "pdf|*.pdf";
                ofd.Multiselect = false;
                if (DialogResult.Cancel == ofd.ShowDialog()) return;

                init();

                if (!File.Exists(ofd.FileName)) return;
                AnalysisInformation(GetStringFromFile(ofd.FileName));
            }
            btnCreate_Click(null, null);
        }

        /// <summary>
        /// 从价格表获取价格字典
        /// </summary>
        private Dictionary<string, int> GetPriceList()
        {
            int index = 0;
            System.Data.DataTable dt = dsConfig.Tables[dropDown.SelectedItemIndex];
            for (int i = 1; i < dt.Columns.Count; i++)
            {
                int v = 0;
                int.TryParse(dt.Columns[i].Caption, out v);
                if(Vehicle <= v)
                {
                    index = i;
                    break;
                }

            }

            Dictionary<string, int> dicPrice = new Dictionary<string, int>();
            foreach (DataRow dr in dt.Rows)
            {
                int p = 0;
                int.TryParse(dr[index].ToString(),out p);
                if (dicPrice.ContainsKey(dr[0].ToString())) continue;
                else
                    dicPrice.Add(dr[0].ToString(),p);
            }
            return dicPrice;
        }
            
        private void btnCreate_Click(object sender, RibbonControlEventArgs e)
        {
            //当前工作表
            Worksheet workSheet = (Worksheet)Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet;
            workSheet.UsedRange.Clear();//清空工作表

            workSheet.Name = "报价单";
            //写入常规信息
            WriteRegularInfo(workSheet);//生成常规信息
            //写入凯恩斯
            WriteCairnsTravelInfo(workSheet);
        }

        #region Excel操作
        private void TitleStyle(Worksheet workSheet,string rX,string rY,string content,bool bBorder,int size)
        {
            Range range = workSheet.get_Range(rX, rY); //获取Excel多个单元格区域：本例做为Excel表头 
            range.Merge(0); //单元格合并动作
            range.Cells[1, 1] = content; //Excel单元格赋值   
            range.Font.Size = size; //设置字体大小   
            range.Font.Underline = false; //设置字体是否有下划线   
            range.Font.Name = "黑体"; //设置字体的种类   
            range.HorizontalAlignment = XlHAlign.xlHAlignCenter; //设置字体在单元格内的对其方式
            if (bBorder)
                range.BorderAround2(1);
        }

        private void NoteStyle(Worksheet workSheet, string rX, string rY, string content, bool bBorder, int size)
        {
            Range range = workSheet.get_Range(rX, rY); //获取Excel多个单元格区域：本例做为Excel表头 
            range.Merge(0); //单元格合并动作
            range.WrapText = true;
            range.Cells[1, 1] = content; //Excel单元格赋值   
            range.Font.Size = size; //设置字体大小   
            range.Font.Underline = false; //设置字体是否有下划线   
            range.HorizontalAlignment = XlHAlign.xlHAlignLeft; //设置字体在单元格内的对其方式
            if (bBorder)
                range.BorderAround2(1);
        }
        void InsertRow(int index)
        {
            Worksheet workSheet = (Worksheet)Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet;
            Microsoft.Office.Interop.Excel.Range row = (Microsoft.Office.Interop.Excel.Range)workSheet.Rows[index, Type.Missing];
            row.EntireRow.Insert(Microsoft.Office.Interop.Excel.XlDirection.xlDown, Microsoft.Office.Interop.Excel.XlInsertFormatOrigin.xlFormatFromLeftOrAbove);
            rowfix += 1;
        }

        void InsertCol(int index)
        {
            Worksheet workSheet = (Worksheet)Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet;
            Microsoft.Office.Interop.Excel.Range col = (Microsoft.Office.Interop.Excel.Range)workSheet.Columns[index, Type.Missing];
            col.EntireColumn.Insert(Microsoft.Office.Interop.Excel.XlDirection.xlToRight, Microsoft.Office.Interop.Excel.XlInsertFormatOrigin.xlFormatFromLeftOrAbove);
        }

        /// <summary>删除行</summary>
        /// <param name="workSheetIndex"></param>
        /// <param name="rowIndex"></param>
        /// <param name="count"></param>
        public void DeleteRows(int rowIndex, int count)
        {
            try
            {
                Worksheet workSheet = (Worksheet)Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet;
                for (int i = 0; i < count; i++)
                {
                    Range range = (Range)workSheet.Rows[rowIndex, Type.Missing];
                    range.Delete(Microsoft.Office.Interop.Excel.XlDirection.xlDown);
                }
                rowfix += count;
            }
            catch (Exception e)
            {
                this.KillExcelProcess();
                throw e;
            }
        }

        /// <summary>插行（在指定行上面插入指定数量行）</summary>
        /// <param name="rowIndex"></param>
        /// <param name="count"></param>
        public void InsertRows(int rowIndex, int count)
        {
            try
            {
                Worksheet workSheet = (Worksheet)Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet;
                Range range = (Microsoft.Office.Interop.Excel.Range)workSheet.Rows[rowIndex];
                for (int i = 0; i < count; i++)
                {
                    range.Insert(Microsoft.Office.Interop.Excel.XlDirection.xlDown);
                }
                rowfix += count;
            }
            catch (Exception e)
            {
                this.KillExcelProcess();
                throw e;
            }
        }

        /// <summary>插列（在指定列右边插入指定数量列）</summary>
        /// <param name="columnIndex"></param>
        /// <param name="count"></param>
        public void InsertCols(int columnIndex, int count)
        {
            try
            {
                Worksheet workSheet = (Worksheet)Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet;
                Range range = (Microsoft.Office.Interop.Excel.Range)(workSheet.Columns[columnIndex]);  //注意：这里和VS的智能提示不一样，第一个参数是columnindex

                for (int i = 0; i < count; i++)
                {
                    range.Insert(Microsoft.Office.Interop.Excel.XlDirection.xlDown);
                }
            }
            catch (Exception e)
            {
                this.KillExcelProcess();
                throw e;
            }
        }

        /// <summary>删除列</summary> 
        /// <param name="columnIndex"></param>
        /// <param name="count"></param>
        public void DeleteColumns(int columnIndex, int count)
        {
            try
            {
                Worksheet workSheet = (Worksheet)Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet;
                for (int i = columnIndex + count - 1; i >= columnIndex; i--)
                {
                    ((Microsoft.Office.Interop.Excel.Range)workSheet.Cells[1, i]).EntireColumn.Delete(0);
                }
            }
            catch (Exception e)
            {
                this.KillExcelProcess();
                throw e;
            }
        }


        /// <summary>复制行（在指定行下面复制指定数量行）</summary>
        /// <param name="rowIndex"></param>
        /// <param name="count"></param>
        public void CopyRows(int rowIndex, int count)
        {
            try
            {
                Worksheet workSheet = (Worksheet)Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet;
                Range range1 = (Microsoft.Office.Interop.Excel.Range)workSheet.Rows[rowIndex];
                for (int i = 1; i <= count; i++)
                {
                    Range range2 = (Microsoft.Office.Interop.Excel.Range)workSheet.Rows[rowIndex + i];
                    range1.Copy(range2);
                }
            }
            catch (Exception e)
            {
                this.KillExcelProcess();
                throw e;
            }
        }

        /// <summary>向单元格写入数据，对当前WorkSheet操作</summary>
        /// <param name="rowIndex">行索引</param>
        /// <param name="columnIndex">列索引</param>
        /// <param name="text">要写入的文本值</param>
        public void SetCells(int rowIndex, int columnIndex, string text)
        {
            try
            {
                Worksheet workSheet = (Worksheet)Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet;
                workSheet.Cells[rowIndex, columnIndex] = text;
            }
            catch
            {
                this.KillExcelProcess();
                throw new Exception("向单元格[" + rowIndex + "," + columnIndex + "]写数据出错！");
            }
        }

        ///// <summary>另存文件</summary>
        //public void SaveAsFile()
        //{
        //    if (this.outputFile == null)
        //        throw new Exception("没有指定输出文件路径！");

        //    try
        //    {
        //        workBook.SaveAs(outputFile, missing, missing, missing, missing, missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, missing, missing, missing, missing, missing);
        //    }
        //    catch (Exception e)
        //    {
        //        throw e;
        //    }
        //    finally
        //    {
        //        this.Quit();
        //    }
        //}

        private void Quit()
        {
            throw new NotImplementedException();
        }


        public void KillExcelProcess()
        {
            Process[] ps = Process.GetProcesses();
            foreach (Process item in ps)
            {
                if (item.ProcessName == "EXCEL")
                {
                    item.Kill();
                }
            }
        }

        public void MoveRangeConent(Range FromRang, Range ToRang)
        {
            FromRang.Copy(ToRang);
            FromRang.Clear();
        }
        #endregion

        private void btnSave_Click(object sender, RibbonControlEventArgs e)
        {
            SaveFileDialog saveDialog = new SaveFileDialog();
            saveDialog.DefaultExt = "pdf";
            saveDialog.Filter = "Pdf文件|*.pdf";
            
            if (DialogResult.OK != saveDialog.ShowDialog()) return; //被点了取消
            Globals.ThisAddIn.Application.ActiveWorkbook.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, saveDialog.FileName);
        }

        private void btnInsertCount_Click(object sender, RibbonControlEventArgs e)
        {
            Range range = (Range)Globals.ThisAddIn.Application.Selection;
            foreach (Range item in range)
            {
                 item.Value = mdConfig.DicConfig["PTC"]["DISC"];
                item.NumberFormatLocal = "0.00%";
            }
        }

        
    }
}
