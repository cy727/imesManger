using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Drawing.Printing;
using Microsoft.Office.Interop.Excel;
using System.Collections;

namespace imesManger
{
    public partial class FormImportData : Form
    {
        private System.Data.SqlClient.SqlConnection sqlConn = new System.Data.SqlClient.SqlConnection();
        private System.Data.SqlClient.SqlCommand sqlComm = new System.Data.SqlClient.SqlCommand();
        private System.Data.SqlClient.SqlDataReader sqldr;
        private System.Data.SqlClient.SqlDataAdapter sqlDA = new System.Data.SqlClient.SqlDataAdapter();
        private System.Data.DataSet dSet = new DataSet();

        public string strConn = "";

        //private ArrayList alIme = new ArrayList();
        public int iRange = 1000000;

        public string sTAG = "Distrubutor";
        public int iSAPControlNumber = 3;
        public string sSAPorder = "SAP order";
        public string sSalesplan = "Sales plan";
        public string sCWIMEI = "CW IMEI";
        public string sMaster = "Master";

        private ArrayList alSAP = new ArrayList();
        private ArrayList alSales = new ArrayList();


        public FormImportData()
        {
            InitializeComponent();
        }

        private void FormImportData_Load(object sender, EventArgs e)
        {
            sqlConn.ConnectionString = strConn;
            sqlComm.Connection = sqlConn;
            sqlDA.SelectCommand = sqlComm;
        }

        private void btnLoad_Click(object sender, EventArgs e)
        {
            int i, j,iTemp=0,iTemp1=0,iTemp2=0,iTemp3=0;
            int iSheet;
            DataSet dsCSV = new DataSet();
            int iI=0,iP=0;
            string sICode="",sPCode="";
            int iTAC = 0;
            string sTAC = "";
            string sNowWeek=ClassIm.GetWeek(dateTimePickerP.Value).ToString();
            int iYear = 0, iWeek=0;



            OpenFileDialog openFileDialogOutput = new OpenFileDialog();
            openFileDialogOutput.Filter = "EXCEL files(*.xlsx)|*.xlsx|EXCEL files(*.xls)|*.xls";//
            openFileDialogOutput.FilterIndex = 0;
            openFileDialogOutput.RestoreDirectory = true;

            if (openFileDialogOutput.ShowDialog() != DialogResult.OK) return;

            Microsoft.Office.Interop.Excel.Application m_xlsApp = null;
            Workbook m_Workbook = null;
            Worksheet m_Worksheet = null;

            toolStripProgressBarALL.Value = 0;
            

            toolStripProgressBarExcel.Value = 0;

            textBoxLOG.Text = "Import Data：\r\n filename:" + openFileDialogOutput.FileName + "\r\n\r\n";
            textBoxERROR.Text = "Error Note：\r\n filename:" + openFileDialogOutput.FileName + "\r\n\r\n";

            sqlConn.Open();

            try
            {
                object objOpt = System.Reflection.Missing.Value;
                m_xlsApp = new Microsoft.Office.Interop.Excel.Application();
                m_Workbook = m_xlsApp.Workbooks.Open(openFileDialogOutput.FileName, objOpt, objOpt, objOpt, objOpt, objOpt, objOpt, objOpt, objOpt, objOpt, objOpt, objOpt, objOpt, objOpt, objOpt);

                toolStripProgressBarALL.Maximum = m_Workbook.Worksheets.Count+2;
                #region accquire
                //TAC Accquire First
                for (iSheet = 1; iSheet <= m_Workbook.Worksheets.Count; iSheet++)
                {

                     m_Worksheet = (Microsoft.Office.Interop.Excel.Worksheet)m_Workbook.Worksheets.get_Item(iSheet);

                    if(m_Worksheet.Name.IndexOf(sMaster)==-1)
                        continue;

                    toolStripProgressBarALL.Value++;
                    textBoxLOG.Text += "\r\nRead TABLE '" + m_Worksheet.Name + "'\r\n";
                                        

                    if (m_Worksheet.UsedRange.Rows.Count < 1 && m_Worksheet.UsedRange.Columns.Count < 8)
                        continue;

                    toolStripProgressBarExcel.Maximum = m_Worksheet.UsedRange.Rows.Count+1;
                    toolStripProgressBarExcel.Value = 0;

                    for (i = 1; i <= m_Worksheet.UsedRange.Rows.Count; i++)
                    {
                        toolStripProgressBarExcel.Value++;
                        toolStripStatusLabelWarn.Text = "LINE:" + toolStripProgressBarExcel.Value.ToString();

                        //跳过空行
                        if (((Range)(m_Worksheet.Cells[i, 1])) == null)
                        {
                            continue;
                        }

                        //跳过描述行
                        if (((Range)(m_Worksheet.Cells[i, 1])).Text.ToString().IndexOf(sTAG) >= 0)
                            continue;

                        //TAC为空
                        if (((Range)(m_Worksheet.Cells[i, 6])).Text.ToString()=="")
                            continue;

                        //iICODE
                        iI=0;sICode="";
                        sqlComm.CommandText = "SELECT ID, [Indentor Name], [Indentor Code] FROM indentor WHERE ([Indentor Code] = N'" + ((Range)(m_Worksheet.Cells[i, 1])).Text.ToString() + "')";
                        sqldr = sqlComm.ExecuteReader();

                        if (sqldr.HasRows)
                        {
                            sqldr.Read();
                            iI = int.Parse(sqldr.GetValue(0).ToString());
                            sICode = sqldr.GetValue(2).ToString();
                            sqldr.Close();
                        }
                        else
                        {
                            textBoxERROR.Text += m_Worksheet.Name + " Line " + i.ToString() + ": NO Indentors Code\r\n";
                            sqldr.Close();
                            continue;
                        }

                        //iPCode
                        iP = 0; sPCode = "";
                        sqlComm.CommandText = "SELECT   ID, [Product Name], [Product Code], [Number of IMEI], Status, [Factory ID], [Factory Code], [Failure Rate] FROM product WHERE   ([Product Code] = N'" + ClassIm.GetPCode(((Range)(m_Worksheet.Cells[i, 2])).Text.ToString()) + "')";
                        sqldr = sqlComm.ExecuteReader();

                        if (sqldr.HasRows)
                        {
                            sqldr.Read();
                            iP = int.Parse(sqldr.GetValue(0).ToString());
                            sPCode = sqldr.GetValue(2).ToString();
                            sqldr.Close();
                        }
                        else
                        {
                            textBoxERROR.Text += m_Worksheet.Name + " Line " + i.ToString() + ": NO Product Code\r\n";
                            sqldr.Close();
                            continue;
                        }

                        //TAC
                        iTAC = 0; sTAC = "";
                        sqlComm.CommandText = "SELECT ID, [TAC Code], [Product ID], [Product Code], [Indentor ID], [Indentor Code], [Init Number] FROM TAC WHERE ([Product Code] = N'"+sPCode+"') AND ([Indentor Code] = N'"+sICode+"')";
                                                sqldr = sqlComm.ExecuteReader();

                        if (sqldr.HasRows) //有TAC
                        {
                            sqldr.Read();
                            iTAC = int.Parse(sqldr.GetValue(0).ToString());
                            sTAC = sqldr.GetValue(1).ToString();

                            sqlComm.CommandText="UPDATE  TAC SET [Init Number] = 0, [Product ID] = "+iP.ToString()+", [Indentor ID] = "+iI.ToString()+", [TAC Code] = N'"+((Range)(m_Worksheet.Cells[i, 6])).Text.ToString()+"' WHERE ([Product Code] = N'"+sPCode+"') AND ([Indentor Code] = N'"+sICode+"')";
                            sqldr.Close();
                            sqlComm.ExecuteNonQuery();
                        }
                        else
                        {
                            sqlComm.CommandText = "INSERT INTO TAC ([Init Number], [Product ID], [Indentor ID], [TAC Code], [Product Code], [Indentor Code]) VALUES (0, " + iP.ToString() + ", " + iI.ToString() + ", N'" + ((Range)(m_Worksheet.Cells[i, 6])).Text.ToString() + "', N'"+sPCode+"', N'"+sICode+"')";
                            sqldr.Close();
                            sqlComm.ExecuteNonQuery();

                            sqlComm.CommandText = "SELECT @@IDENTITY";
                            sqldr = sqlComm.ExecuteReader();
                            sqldr.Read();
                            iTAC = Convert.ToInt32(sqldr.GetValue(0).ToString());
                            sqldr.Close();
                            sTAC = ((Range)(m_Worksheet.Cells[i, 6])).Text.ToString();
                        }



                        //accquire
                        if (((Range)(m_Worksheet.Cells[i, 7])).Text.ToString() == "")
                            continue;
                        if (((Range)(m_Worksheet.Cells[i, 8])).Text.ToString() == "")
                            continue;

                        sqlComm.CommandText = "SELECT ID, [Product ID], [Product Code], [Indentor ID], [Indentor Code], [Start Number], [End Number], Date, Year, [Num of Week],  [TAC ID], [TAC Code], Status FROM acquire WHERE ([Product Code] = N'" + sPCode + "') AND ([Indentor Code] = N'" + sICode + "')";
                        sqldr = sqlComm.ExecuteReader();

                        if (sqldr.HasRows) //有申请
                        {
                            sqlComm.CommandText = "UPDATE  acquire SET [Start Number] = " + ((Range)(m_Worksheet.Cells[i, 7])).Text.ToString() + ", [End Number] = " + ((Range)(m_Worksheet.Cells[i, 8])).Text.ToString() + ", Date = CONVERT(DATETIME, '" + dateTimePickerP.Value.ToShortDateString() + " 00:00:00', 102), Year = " + dateTimePickerP.Value.Year.ToString() + ",  [Num of Week] = N'" + ClassIm.GetWeek(dateTimePickerP.Value).ToString() + "', [Product ID] = " + iP.ToString() + ", [Indentor ID] = " + iI.ToString() + ", [TAC ID] = " + iTAC.ToString() + ", [TAC Code] = N'" + sTAC + "' WHERE ([Product Code] = N'" + sPCode + "') AND ([Indentor Code] = N'" + sICode + "')"; ;
                        }
                        else
                        {
                            sqlComm.CommandText = "INSERT INTO acquire ([Start Number], [End Number], Date, Year, [Num of Week], [Product ID], [Indentor ID], [TAC ID], [TAC Code], [Product Code], [Indentor Code]) VALUES (" + ((Range)(m_Worksheet.Cells[i, 7])).Text.ToString() + ", " + ((Range)(m_Worksheet.Cells[i, 8])).Text.ToString() + ", CONVERT(DATETIME, '" + dateTimePickerP.Value.ToShortDateString() + " 00:00:00', 102), " + dateTimePickerP.Value.Year.ToString() + ", N'" + ClassIm.GetWeek(dateTimePickerP.Value).ToString() + "', " + iP.ToString() + ", "+iI.ToString()+", "+iTAC.ToString()+", N'"+sTAC+"', N'"+sPCode+"', N'"+sICode+"')";
                        }
                        sqldr.Close();
                        sqlComm.ExecuteNonQuery();

                    }

                }
                #endregion

                //TAC
                sqlComm.CommandText = "SELECT ID, [TAC Code], [Product ID], [Product Code], [Indentor ID], [Indentor Code], [Init Number] FROM      TAC";
                if (dSet.Tables.Contains("TAC")) dSet.Tables["TAC"].Clear();
                sqlDA.Fill(dSet, "TAC");

                //manufacture & Order
                alSales.Clear();
                alSAP.Clear();
                #region Order
                for (iSheet = 1; iSheet <= m_Workbook.Worksheets.Count; iSheet++)
                {
                    m_Worksheet = (Microsoft.Office.Interop.Excel.Worksheet)m_Workbook.Worksheets.get_Item(iSheet);

                    //if (m_Worksheet.Name.IndexOf(sSAPorder) >= 0)
                    //    continue;
                    //if (m_Worksheet.Name.IndexOf(sSalesplan) >= 0)
                    //    continue;

                    //**************************************************************************************
                    //manufacture
                    if (m_Worksheet.Name.IndexOf(sCWIMEI) >= 0)
                    {
                        toolStripProgressBarALL.Value++;
                        textBoxLOG.Text += "\r\nRead TABLE '" + m_Worksheet.Name + "'\r\n";

                        if (m_Worksheet.UsedRange.Rows.Count < 2 && m_Worksheet.UsedRange.Columns.Count < 4)
                        {
                            textBoxERROR.Text += m_Worksheet.Name+" DATA Format Wrong\r\n";
                            continue;
                        }

                        toolStripProgressBarExcel.Maximum = m_Worksheet.UsedRange.Rows.Count + 1;
                        toolStripProgressBarExcel.Value = 0;
                        for (i = 2; i <= m_Worksheet.UsedRange.Rows.Count; i++)
                        {
                            toolStripProgressBarExcel.Value++;
                            toolStripStatusLabelWarn.Text = "LINE:" + toolStripProgressBarExcel.Value.ToString();

                            //if (i == 51)
                            //{
                            //    int kkk = 0;
                            //}
                            //跳过空行
                            if (((Range)(m_Worksheet.Cells[i, 3])) == null)
                            {
                                continue;
                            }

                            if (((Range)(m_Worksheet.Cells[i, 3])).Value2.ToString().Length != 15)
                            {
                                textBoxERROR.Text += m_Worksheet.Name + " Line " + i.ToString() + ": Data Error - Length is not 15\r\n";
                                continue;
                            }

                            //iICODE
                            iI = 0; sICode = "";
                            sqlComm.CommandText = "SELECT ID, [Indentor Name], [Indentor Code] FROM indentor WHERE ([Indentor Code] = N'" + ((Range)(m_Worksheet.Cells[i, 1])).Text.ToString() + "')";
                            sqldr = sqlComm.ExecuteReader();

                            if (sqldr.HasRows)
                            {
                                sqldr.Read();
                                iI = int.Parse(sqldr.GetValue(0).ToString());
                                sICode = sqldr.GetValue(2).ToString();
                                sqldr.Close();
                            }
                            else
                            {
                                textBoxERROR.Text += m_Worksheet.Name + " Line " + i.ToString() + ": NO Indentors Code\r\n";
                                sqldr.Close();
                                continue;
                            }

                            //iPCode
                            iP = 0; sPCode = "";
                            sqlComm.CommandText = "SELECT   ID, [Product Name], [Product Code], [Number of IMEI], Status, [Factory ID], [Factory Code], [Failure Rate] FROM product WHERE   ([Product Code] = N'" + ClassIm.GetPCode(((Range)(m_Worksheet.Cells[i, 2])).Text.ToString()) + "')";
                            sqldr = sqlComm.ExecuteReader();

                            if (sqldr.HasRows)
                            {
                                sqldr.Read();
                                iP = int.Parse(sqldr.GetValue(0).ToString());
                                sPCode = sqldr.GetValue(2).ToString();
                                sqldr.Close();
                            }
                            else
                            {
                                textBoxERROR.Text += m_Worksheet.Name + " Line " + i.ToString() + ": NO Product Code\r\n";
                                sqldr.Close();
                                continue;
                            }

                            //检查TAC
                            iTAC = 0; sTAC = "";
                            var qTAC = from dt2 in dSet.Tables["TAC"].AsEnumerable()//得到TAC初始值
                                       where (dt2.Field<string>("Product Code") == sPCode) && (dt2.Field<string>("Indentor Code") == sICode)//条件
                                       select dt2;
                            if (qTAC.Count() <= 0)
                            {
                                continue;
                            }
                            foreach (var itemTAC in qTAC)//显示查询结果//有TAC
                            {
                                iTAC = itemTAC.Field<int>("Init Number");
                                sTAC = itemTAC.Field<string>("TAC Code");
                                break;
                            }

                            
                            //manufacture
                            iTemp1 = int.Parse(((Range)(m_Worksheet.Cells[i, 3])).Value2.ToString().Substring(8, 6));
                            if (((Range)(m_Worksheet.Cells[i, 3])).Value2.ToString().Substring(0, 8) == "--------")
                            {
                                iTemp1 = -1*iTemp1;
                            }


                            sqlComm.CommandText = "SELECT ID, [Product ID], [Product Code], [Indentor ID], [Indentor Code], [Start Number], [End Number], [Acquire ID], Date, Year, [Num of Week], [TAC ID], [TAC Code], [Total Number] FROM actual WHERE ([Product Code] = N'" + sPCode + "') AND ([Indentor Code] = N'" + sICode + "')";
                            sqldr = sqlComm.ExecuteReader();

                            if (sqldr.HasRows) 
                            {
                                sqldr.Read();
                                iTemp = 0;
                                if (sqldr.GetValue(6).ToString() == "")
                                    iTemp = 0;
                                else
                                    iTemp=int.Parse(sqldr.GetValue(6).ToString())+1;

                                sqlComm.CommandText = "UPDATE  actual SET [Product ID] = " + iP.ToString() + ", [Indentor ID] = " + iI.ToString() + ", [Start Number] = " + iTemp.ToString() + ", [End Number] = " + iTemp1.ToString() + ", Date = CONVERT(DATETIME,  '" + dateTimePickerP.Value.ToShortDateString() + " 00:00:00', 102), Year = " + dateTimePickerP.Value.Year.ToString() + ", [Num of Week] = N'" + sNowWeek + "', [TAC ID] = "+iTAC.ToString()+", [TAC Code] = N'"+sTAC+"', [Total Number] = "+iTemp1.ToString()+" WHERE ([Product Code] = N'" + sPCode + "') AND ([Indentor Code] = N'" + sICode + "')";
                            }
                            else
                            {
                                sqlComm.CommandText = "INSERT INTO actual ([Product ID], [Indentor ID], [Start Number], [End Number], Date, Year, [Num of Week], [TAC ID], [TAC Code],  [Total Number], [Product Code], [Indentor Code]) VALUES (" + iP.ToString() + ", " + iI.ToString() + ", 0, " + iTemp1 + ", CONVERT(DATETIME, '" + dateTimePickerP.Value.ToShortDateString() + " 00:00:00', 102), "+dateTimePickerP.Value.Year.ToString()+", N'"+sNowWeek+"', "+iTAC.ToString()+", N'"+sTAC+"', "+iTemp1.ToString()+", N'"+sPCode+"', N'"+sICode+"')";
                            }
                            sqldr.Close();
                            sqlComm.ExecuteNonQuery();

                            //backlog
                            if (((Range)(m_Worksheet.Cells[i, 4])).Text.ToString() == "")
                                iTemp1 = 0;
                            else
                                iTemp1 = int.Parse(((Range)(m_Worksheet.Cells[i, 4])).Value2.ToString());

                            sqlComm.CommandText = "SELECT ID, [Product ID], [Product Code], [Indentor ID], [Indentor Code], [Start Number], [End Number], [Acquire ID], Date, Year, [Num of Week], Numbers, [Order Start Number], [Order End Number], backlog FROM orders WHERE ([Product Code] = N'" + sPCode + "') AND ([Indentor Code] = N'" + sICode + "') AND (Year = "+dateTimePickerP.Value.Year.ToString()+") AND ([Num of Week] = N'"+sNowWeek+"')";
                            sqldr = sqlComm.ExecuteReader();
                            if (sqldr.HasRows)
                            {
                                sqlComm.CommandText = "UPDATE  orders SET backlog = "+iTemp1.ToString()+" WHERE ([Product Code] = N'" + sPCode + "') AND ([Indentor Code] = N'" + sICode + "') AND (Year = "+dateTimePickerP.Value.Year.ToString()+") AND ([Num of Week] = N'"+sNowWeek+"')";
                            }
                            else
                            {
                                sqlComm.CommandText = "INSERT INTO orders ([Product ID], [Product Code], [Indentor ID], [Indentor Code], [Start Number], [End Number], Date, Year, [Num of Week],  Numbers, [Order Start Number], [Order End Number], backlog) VALUES   ("+iP.ToString()+", N'"+sPCode+"', "+iI.ToString()+", N'"+sICode+"', 0, 0, CONVERT(DATETIME, '"+dateTimePickerP.Value.ToShortDateString()+" 00:00:00', 102), "+dateTimePickerP.Value.Year.ToString()+", N'"+sNowWeek+"', 0, 0, 0, "+iTemp1+")";
                            }
                            sqldr.Close();
                            sqlComm.ExecuteNonQuery();

                        }

                    }


                    //****************************************************************************************
                    //SAP
                    #region SAP
                    if (m_Worksheet.Name.IndexOf(sSAPorder) >= 0)
                    {
                        toolStripProgressBarALL.Value++;
                        textBoxLOG.Text += "\r\nRead TABLE '" + m_Worksheet.Name + "'\r\n";


                        if (m_Worksheet.UsedRange.Rows.Count < 2 && m_Worksheet.UsedRange.Columns.Count < 14)
                        {
                            textBoxERROR.Text += m_Worksheet.Name + " DATA Format Wrong\r\n";
                            continue;
                        }

                        toolStripProgressBarExcel.Maximum = m_Worksheet.UsedRange.Rows.Count + 1;
                        toolStripProgressBarExcel.Value = 0;
                        for (i = 2; i <= m_Worksheet.UsedRange.Rows.Count; i++)
                        {
                            //if (i == 553)
                            //{
                            //    int kkk = 0;
                            //}
                            //跳过空行
                            toolStripProgressBarExcel.Value++;
                            toolStripStatusLabelWarn.Text = "LINE:" + toolStripProgressBarExcel.Value.ToString();
                            if (((Range)(m_Worksheet.Cells[i, 14])) == null)
                            {
                                break;
                            }

                            if (((Range)(m_Worksheet.Cells[i, 14])).Text.ToString() == "")
                            {
                                break;
                            }

                            if (((Range)(m_Worksheet.Cells[i, 1])).Value2.ToString() != "") //第一列检查
                            {
                                continue;
                            }

                            sPCode = ClassIm.GetPCode(((Range)(m_Worksheet.Cells[i, 9])).Value2.ToString());
                            sICode = ((Range)(m_Worksheet.Cells[i, 5])).Value2.ToString();
                            string[] sT = ((Range)(m_Worksheet.Cells[i, 11])).Value2.ToString().Split('.');
                            var alQuery = from cImeOrder cime in alSAP
                                          where (cime.sPCode == sPCode && (cime.sICode == sICode) && (cime.iYear == int.Parse(sT[0])) && (cime.iWeek == int.Parse(sT[1])))
                                          select cime;

                            if (alQuery.Count() <= 0) //没有记录
                            {
                                //iICODE
                                iI = 0;
                                sqlComm.CommandText = "SELECT ID, [Indentor Name], [Indentor Code] FROM indentor WHERE ([Indentor Code] = N'" + sICode + "')";
                                sqldr = sqlComm.ExecuteReader();

                                if (sqldr.HasRows)
                                {
                                    sqldr.Read();
                                    iI = int.Parse(sqldr.GetValue(0).ToString());
                                    sICode = sqldr.GetValue(2).ToString();
                                    sqldr.Close();
                                }
                                else
                                {
                                    textBoxERROR.Text += m_Worksheet.Name + " Line " + i.ToString() + ": NO Indentors Code\r\n";
                                    sqldr.Close();
                                    continue;
                                }

                                //iPCode
                                iP = 0;
                                sqlComm.CommandText = "SELECT   ID, [Product Name], [Product Code], [Number of IMEI], Status, [Factory ID], [Factory Code], [Failure Rate] FROM product WHERE   ([Product Code] = N'" + sPCode + "')";
                                sqldr = sqlComm.ExecuteReader();

                                if (sqldr.HasRows)
                                {
                                    sqldr.Read();
                                    iP = int.Parse(sqldr.GetValue(0).ToString());
                                    sPCode = sqldr.GetValue(2).ToString();
                                    sqldr.Close();
                                }
                                else
                                {
                                    textBoxERROR.Text += m_Worksheet.Name + " Line " + i.ToString() + ": NO Product Code\r\n";
                                    sqldr.Close();
                                    continue;
                                }

                                cImeOrder ci = new cImeOrder(iP, sPCode, iI, sICode, int.Parse(((Range)(m_Worksheet.Cells[i, 14])).Value2.ToString()), int.Parse(sT[0]), int.Parse(sT[1]));
                                alSAP.Add(ci);
                            }
                            else //存在记录，直接加入
                            {
                                foreach (cImeOrder s in alQuery)
                                {
                                    s.iNum += int.Parse(((Range)(m_Worksheet.Cells[i, 14])).Value2.ToString());
                                }
                            }
                        }
                    }
                    #endregion

                    //****************************************************************************************
                    //SALES 
                    #region SALES
                    iTemp=int.Parse(ClassIm.GetWeek(dateTimePickerP.Value,iSAPControlNumber-1));
                    iTemp=int.Parse(ClassIm.sYear)*100+iTemp;
                    if (m_Worksheet.Name.IndexOf(sSalesplan) >= 0)
                    {
                        toolStripProgressBarALL.Value++;
                        textBoxLOG.Text += "\r\nRead TABLE '" + m_Worksheet.Name + "'\r\n";

                        if (m_Worksheet.UsedRange.Rows.Count < 2 && m_Worksheet.UsedRange.Columns.Count < 3)
                        {
                            textBoxERROR.Text += m_Worksheet.Name + " DATA Format Wrong\r\n";
                            continue;
                        }

                        toolStripProgressBarExcel.Maximum = m_Worksheet.UsedRange.Rows.Count + 1;
                        toolStripProgressBarExcel.Value = 0;
                        for (i = 2; i <= m_Worksheet.UsedRange.Rows.Count; i++)
                        {
                            toolStripProgressBarExcel.Value++;
                            toolStripStatusLabelWarn.Text = "LINE:" + toolStripProgressBarExcel.Value.ToString();
                            //if (i == 96)
                            //{
                            //    int kkk = 0;
                            //}
                            //iICODE
                            iI = 0; sICode = "";
                            sqlComm.CommandText = "SELECT ID, [Indentor Name], [Indentor Code] FROM indentor WHERE ([Indentor Code] = N'" + ((Range)(m_Worksheet.Cells[i, 1])).Value2.ToString() + "')";
                            sqldr = sqlComm.ExecuteReader();

                            if (sqldr.HasRows)
                            {
                                sqldr.Read();
                                iI = int.Parse(sqldr.GetValue(0).ToString());
                                sICode = sqldr.GetValue(2).ToString();
                                sqldr.Close();
                            }
                            else
                            {
                                textBoxERROR.Text += m_Worksheet.Name + " Line " + i.ToString() + ": NO Indentors Code\r\n";
                                sqldr.Close();
                                continue;
                            }

                            //iPCode
                            iP = 0; sPCode = "";
                            sqlComm.CommandText = "SELECT   ID, [Product Name], [Product Code], [Number of IMEI], Status, [Factory ID], [Factory Code], [Failure Rate] FROM product WHERE   ([Product Code] = N'" + ClassIm.GetPCode(((Range)(m_Worksheet.Cells[i, 2])).Value2.ToString()) + "')";
                            sqldr = sqlComm.ExecuteReader();

                            if (sqldr.HasRows)
                            {
                                sqldr.Read();
                                iP = int.Parse(sqldr.GetValue(0).ToString());
                                sPCode = sqldr.GetValue(2).ToString();
                                sqldr.Close();
                            }
                            else
                            {
                                textBoxERROR.Text += m_Worksheet.Name + " Line " + i.ToString() + ": NO Product Code\r\n";
                                sqldr.Close();
                                continue;
                            }

                            for (j = 3; j <= m_Worksheet.UsedRange.Columns.Count; j++)
                            {
                                //跳过空行

                                if (((Range)(m_Worksheet.Cells[i, j])) == null)
                                {
                                    continue;
                                }

                                if (((Range)(m_Worksheet.Cells[1, j])).Text.ToString() == "")
                                {
                                    continue;
                                }
                                if (((Range)(m_Worksheet.Cells[1, j])).Text.ToString().Length < 8)
                                {
                                    continue;
                                }

                                iYear = int.Parse(((Range)(m_Worksheet.Cells[1, j])).Text.ToString().Substring(0, 4));
                                iWeek = int.Parse(((Range)(m_Worksheet.Cells[1, j])).Text.ToString().Substring(6, 2));


                                if (iYear * 100 + iWeek <= iTemp)
                                {
                                    if (checkBoxDetail.Checked)
                                    {
                                        textBoxLOG.Text += "Sale plan" + iYear.ToString() + "-" + iWeek.ToString() + " ignore \r\n";
                                    }
                                    iTemp3 = 0;
                                }
                                else
                                    iTemp3 = int.Parse(((Range)(m_Worksheet.Cells[i, j])).Value2.ToString());

                                var alQuery = from cImeOrder cime in alSales
                                              where ((cime.sPCode == sPCode) && (cime.sICode == sICode) && (cime.iYear == iYear) && (cime.iWeek == iWeek))
                                              select cime;

                                if (alQuery.Count() <= 0) //没有记录
                                {
                                    cImeOrder ci = new cImeOrder(iP, sPCode, iI, sICode, iTemp3, iYear, iWeek);
                                    alSales.Add(ci);
                                }
                                else //存在记录，直接加入
                                {
                                    foreach (cImeOrder s in alQuery)
                                    {
                                        s.iNum = iTemp3;
                                    }
                                }

                            }

                        }
                    }
                    #endregion


                }//read EXCEL


                //处理ORDER
                //SAP有效周记数
                toolStripProgressBarALL.Value = toolStripProgressBarALL.Maximum - 2;
                textBoxLOG.Text += "\r\nCalculate SAP order & Sales plan:\r\n";

                iTemp=int.Parse(ClassIm.GetWeek(dateTimePickerP.Value,iSAPControlNumber-1));
                iTemp=int.Parse(ClassIm.sYear)*100+iTemp;

                iTemp1=dateTimePickerP.Value.Year*100+int.Parse(sNowWeek);

                toolStripProgressBarExcel.Maximum = alSAP.Count + 1;
                toolStripProgressBarExcel.Value = 0;
                foreach (cImeOrder cime in alSAP)
                {
                    toolStripProgressBarExcel.Value++;
                    toolStripStatusLabelWarn.Text = "ITEM:" + toolStripProgressBarExcel.Value.ToString();
                    this.Refresh();
                    textBoxLOG.ScrollToCaret(); textBoxERROR.ScrollToCaret();

                    var alQuery = from cImeOrder cimeSales in alSales
                                  where ((cimeSales.sPCode == cime.sPCode) && (cimeSales.sICode == cime.sICode) && (cimeSales.iYear == cime.iYear) && (cimeSales.iWeek == cime.iWeek))
                                  select cimeSales;

                    if (alQuery.Count() < 1) //没有对应的Sales
                    {
                        alSales.Add(cime);
                    }
                    else
                    {
                        foreach (cImeOrder ci in alQuery) //存在对应的sales记录 ci->sales
                        {
                            if (cime.iYear * 100 + cime.iWeek > iTemp) //不在SAP控制范围内
                            {
                                ci.iNum = Math.Max(ci.iNum, cime.iNum); //取大值

                                if (checkBoxDetail.Checked)
                                {
                                    textBoxLOG.Text += ci.iYear.ToString()+"-"+ci.iWeek.ToString()+" "+ci.sICode + "|" + ci.sPCode + ":" + ci.iNum.ToString() + "=MAX(" + ci.iNum + "," + cime.iNum + ")\r\n";
                                }

                            }
                            else //在sap控制范围内
                            {
                                ci.iNum = cime.iNum; //取SAP值
                                if (checkBoxDetail.Checked)
                                {
                                    textBoxLOG.Text += ci.iYear.ToString() + "-" + ci.iWeek.ToString() + " " + ci.sICode + "|" + ci.sPCode + ":SAP(" + ci.iNum + ")\r\n";
                                }
                            }
                        }
                    }

                }

                //导入数据库
                System.Data.SqlClient.SqlTransaction sqlta;
                sqlta = sqlConn.BeginTransaction();
                sqlComm.Transaction = sqlta;

                toolStripProgressBarALL.Value = toolStripProgressBarALL.Maximum - 1;
                textBoxLOG.Text += "\r\nInsert Data to Database:\r\n";
                toolStripProgressBarExcel.Maximum = alSales.Count + 1;
                toolStripProgressBarExcel.Value = 0;
                try

                {
                    foreach (cImeOrder cime1 in alSales)
                    {
                        toolStripProgressBarExcel.Value++;
                        ScrollToCaret();
                        this.Refresh();



                        toolStripStatusLabelWarn.Text = "ITEM:" + toolStripProgressBarExcel.Value.ToString();
                        sqlComm.CommandText = "SELECT ID, [Product ID], [Product Code], [Indentor ID], [Indentor Code], [Start Number], [End Number], [Acquire ID], Date, Year, [Num of Week], Numbers, [Order Start Number], [Order End Number], backlog FROM orders WHERE ([Product Code] = N'" + cime1.sPCode + "') AND ([Indentor Code] = N'" + cime1.sICode + "') AND (Year = " + cime1.iYear.ToString() + ") AND ([Num of Week] = N'" + cime1.iWeek.ToString() + "')";
                        sqldr = sqlComm.ExecuteReader();
                        if (sqldr.HasRows)
                        {
                            sqlComm.CommandText = "UPDATE  orders SET backlog = " + iTemp1.ToString() + " WHERE ([Product Code] = N'" + sPCode + "') AND ([Indentor Code] = N'" + sICode + "') AND (Year = " + dateTimePickerP.Value.Year.ToString() + ") AND ([Num of Week] = N'" + sNowWeek + "')";

                            sqlComm.CommandText = "UPDATE  orders SET [Start Number] = 0, [End Number] = 0, [Acquire ID] = 0, Numbers = " + cime1.iNum.ToString() + ", [Order Start Number] =0, [Order End Number] =0 WHERE ([Product Code] = N'" + cime1.sPCode + "') AND ([Indentor Code] = N'" + cime1.sICode + "') AND (Year = " + cime1.iYear.ToString() + ") AND ([Num of Week] = N'" + cime1.iWeek.ToString() + "')";
                        }
                        else
                        {
                            sqlComm.CommandText = "INSERT INTO orders ([Product ID], [Product Code], [Indentor ID], [Indentor Code], [Start Number], [End Number], [Acquire ID], Date, Year,  [Num of Week], Numbers,[Order Start Number],[Order End Number],backlog) VALUES  (" + cime1.iPCode.ToString() + ", N'" + cime1.sPCode + "', " + cime1.iICode.ToString() + ", N'" + cime1.sICode + "', 0, 0, 0, CONVERT(DATETIME, '" + System.DateTime.Now.ToShortDateString() + " 00:00:00', 102), " + cime1.iYear.ToString() + ", N'" + cime1.iWeek.ToString() + "', " + cime1.iNum.ToString() + ",0,0,0)";

                        }
                        sqldr.Close();
                        sqlComm.ExecuteNonQuery();
                    }
                    sqlta.Commit();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("error：" + ex.Message.ToString(), "error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    sqlta.Rollback();
                }
                finally
                {
                   
                }

                toolStripProgressBarExcel.Value = toolStripProgressBarExcel.Maximum;
                toolStripProgressBarALL.Value = toolStripProgressBarALL.Maximum;
                toolStripStatusLabelWarn.Text = "";

                #endregion

                MessageBox.Show("Input Finish", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception exc)
            {
                MessageBox.Show("Input Fail："+exc.Message.ToString(), "Error");
            }
            finally
            {
                sqlConn.Close();
                m_Worksheet = null;
                m_Workbook = null;
                m_xlsApp.Quit();
                int generation = System.GC.GetGeneration(m_xlsApp);
                m_xlsApp = null;
                System.GC.Collect(generation);
            }
        }

        private void btnLog_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveFileDialogOutput = new SaveFileDialog();
            saveFileDialogOutput.Filter = "Txt files(*.txt)|*.txt";//
            saveFileDialogOutput.FilterIndex = 0;
            saveFileDialogOutput.RestoreDirectory = true;
            saveFileDialogOutput.FileName = "LOG";

            if (saveFileDialogOutput.ShowDialog() != DialogResult.OK) return;

            FileStream fs = File.Create(saveFileDialogOutput.FileName);
            StreamWriter sw = new StreamWriter(fs);
            sw.WriteLine(textBoxLOG.Text);
            sw.Close();
            fs.Close();
        }

        private void btnError_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveFileDialogOutput = new SaveFileDialog();
            saveFileDialogOutput.Filter = "Txt files(*.txt)|*.txt";//
            saveFileDialogOutput.FilterIndex = 0;
            saveFileDialogOutput.RestoreDirectory = true;
            saveFileDialogOutput.FileName = "ERROR";

            if (saveFileDialogOutput.ShowDialog() != DialogResult.OK) return;

            FileStream fs = File.Create(saveFileDialogOutput.FileName);
            StreamWriter sw = new StreamWriter(fs);
            sw.WriteLine(textBoxERROR.Text);
            sw.Close();
            fs.Close();

        }

        private void ScrollToCaret()
        {
            this.textBoxLOG.ScrollToCaret();
            this.textBoxLOG.Focus();
            this.textBoxLOG.Select(this.textBoxLOG.TextLength, 0);
            this.textBoxLOG.ScrollToCaret();
        }
    }

    class cImeOrder
    {
        public int iPCode = 0;
        public int iICode = 0;

        public string sPCode = "";
        public string sICode = "";

        public int iNum = 0;

        public int iYear = 0;
        public int iWeek = 0;

        public cImeOrder(int iPCode, string sPCode, int iICode,string sICode, int iNum,int iYear,int iWeek) //带参数的构造函数
        {
            this.iPCode = iPCode;
            this.iICode = iICode;
            this.sPCode = sPCode;
            this.sICode = sICode;
            this.iNum = iNum;

            this.iYear = iYear;
            this.iWeek = iWeek;
        }
    }
}
