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
using System.Threading;

namespace imesManger
{
    public partial class FormImportDataOffLine : Form
    {
        //private System.Data.SqlClient.SqlConnection sqlConn = new System.Data.SqlClient.SqlConnection();
        //private System.Data.SqlClient.SqlCommand sqlComm = new System.Data.SqlClient.SqlCommand();
        //private System.Data.SqlClient.SqlDataReader sqldr;
        //private System.Data.SqlClient.SqlDataAdapter sqlDA = new System.Data.SqlClient.SqlDataAdapter();
        //private System.Data.DataSet dSet = new DataSet();

        public string strConn = "";

        //private ArrayList alIme = new ArrayList();
        public int iRange = 1000000;

        //product table
        public string sProduct = "Reserve IMEI Info";
        private int iProduct = -1;
        private string[] sTableProduct = { "Product hierarchy", "Model number", "DS", "Failure Rate", "Factory","Type Designation" };
        private int[] iTableProduct = { -1,-1,-1,-1,-1,-1};

        //Customer
        public string sCustomer = "Customer";
        private int iCustomer = -1;
        private string[] sTableCustomer = { "Customer", "Customer Code" };
        private int[] iTableCustomer = { -1, -1 };

        //TAC
        public string sTACNumber = "Reserve IMEI Info";
        private int iTACNumber = -1;
        private string[] sTableTAC = { "TAC", "Product hierarchy", "Distrubutor" };
        private int[] iTableTAC = { -1, -1, -1 };

        //Acquire
        public string sAcquire = "Reserve IMEI Info";
        private int iAcquire = -1;
        private string[] sTableAcquire = { "TAC", "Product hierarchy", "Distrubutor","SN_" };
        private int[] iTableAcquire = { -1, -1, -1, -1 };
        List<ClassIndexofOrder> listIndex = new List<ClassIndexofOrder>();

        //Actual,backlog
        public string sActual = "CW IMEI+backlog";
        private int iActual = -1;
        private string[] sTableActual = { "Type Designation", "Profile", "Last IMEI#", "backlog" };
        private int[] iTableActual = { -1, -1, -1, -1 };

        //Order
        public string sSAPorder = "SAP order";
        public string sSalesplan = "Sales plan";
        private int iSAPorder = -1;
        private int iSalesplan = -1;
        private string[] sTableSAPorder = { "Product hierarchy", "Sold-to pt", "Nokia GI week", "Confirmed Qty" };
        private int[] iTableSAPorder = { -1, -1, -1, -1 };
        private string[] sTableSalesplan = { "Product hierarchy", "Distrubutor" };
        private int[] iTableSalesplan = { -1, -1 };


        public int iSAPControlNumber = 3;
        
        //public string sMaster = "Master";


        private ArrayList alSAP = new ArrayList();
        private ArrayList alSales = new ArrayList();

        public System.Data.DataTable dtProduct = new System.Data.DataTable();
        public System.Data.DataTable dtCustumor = new System.Data.DataTable();
        public System.Data.DataTable dtTAC = new System.Data.DataTable();
        public System.Data.DataTable dtAcquire = new System.Data.DataTable();
        public System.Data.DataTable dtActual = new System.Data.DataTable();
        public System.Data.DataTable dtOrder = new System.Data.DataTable();
        public System.Data.DataTable dtBacklog = new System.Data.DataTable();

        public Form f;

        public int iNumofWeek = 11;
        public int iNumofPass = 1;

        Thread threadUI = null;
        Thread threadWORK = null;

        public string ExcelPathName = "";

        private int toolStripProgressBarALLMaximum = 0;
        private int toolStripProgressBarALLValue = 0;
        private int toolStripProgressBarExcelMaximum = 0;
        private int toolStripProgressBarExcelValue = 0;
        

        private string textBoxERRORText="";
        private string textBoxLOGText = "";
        private string toolStripStatusLabelWarnText = "";



        public FormImportDataOffLine()
        {
            InitializeComponent();
        }

        private void FormImportDataOffLine_Load(object sender, EventArgs e)
        {
            dtProduct.Columns.Add("ID", System.Type.GetType("System.Decimal"));//0
            dtProduct.Columns.Add("Product Code", System.Type.GetType("System.String"));
            dtProduct.Columns.Add("Product Name", System.Type.GetType("System.String"));
            dtProduct.Columns.Add("Number of IMEI", System.Type.GetType("System.Decimal"));
            dtProduct.Columns.Add("Failure Rate", System.Type.GetType("System.Decimal"));
            dtProduct.Columns.Add("Factory", System.Type.GetType("System.String"));
            dtProduct.Columns.Add("Type Designation", System.Type.GetType("System.String"));

            dtCustumor.Columns.Add("ID", System.Type.GetType("System.Decimal"));//0
            dtCustumor.Columns.Add("Customer Code", System.Type.GetType("System.String"));
            dtCustumor.Columns.Add("Customer Name", System.Type.GetType("System.String"));

            dtTAC.Columns.Add("ID", System.Type.GetType("System.Decimal"));//0
            dtTAC.Columns.Add("TAC Code", System.Type.GetType("System.String"));
            dtTAC.Columns.Add("Product ID", System.Type.GetType("System.Decimal"));
            dtTAC.Columns.Add("Product Code", System.Type.GetType("System.String"));
            dtTAC.Columns.Add("Customer ID", System.Type.GetType("System.Decimal"));
            dtTAC.Columns.Add("Customer Code", System.Type.GetType("System.String"));
            dtTAC.Columns.Add("Init Number", System.Type.GetType("System.Decimal"));

            dtAcquire.Columns.Add("ID", System.Type.GetType("System.Decimal"));//0
            dtAcquire.Columns.Add("Product ID", System.Type.GetType("System.Decimal"));
            dtAcquire.Columns.Add("Product Code", System.Type.GetType("System.String"));
            dtAcquire.Columns.Add("Customer ID", System.Type.GetType("System.Decimal"));
            dtAcquire.Columns.Add("Customer Code", System.Type.GetType("System.String"));
            dtAcquire.Columns.Add("End Number", System.Type.GetType("System.Decimal"));
            dtAcquire.Columns.Add("Year", System.Type.GetType("System.Decimal"));
            dtAcquire.Columns.Add("Num of Week", System.Type.GetType("System.String"));
            dtAcquire.Columns.Add("Start Number", System.Type.GetType("System.Decimal"));
            dtAcquire.Columns.Add("Date", System.Type.GetType("System.String"));
            dtAcquire.Columns.Add("TAC ID", System.Type.GetType("System.Decimal"));
            dtAcquire.Columns.Add("TAC Code", System.Type.GetType("System.String"));
            dtAcquire.Columns.Add("Init Number", System.Type.GetType("System.Decimal"));
            dtAcquire.Columns.Add("Status", System.Type.GetType("System.Decimal"));

            dtActual.Columns.Add("Product ID", System.Type.GetType("System.Decimal"));//0
            dtActual.Columns.Add("Product Code", System.Type.GetType("System.String"));//0
            dtActual.Columns.Add("Customer ID", System.Type.GetType("System.Decimal"));//0
            dtActual.Columns.Add("Customer Code", System.Type.GetType("System.String"));//0
            dtActual.Columns.Add("Start Number", System.Type.GetType("System.Decimal"));//0
            dtActual.Columns.Add("End Number", System.Type.GetType("System.Decimal"));//0
            dtActual.Columns.Add("Acquire ID", System.Type.GetType("System.Decimal"));//0
            dtActual.Columns.Add("Date", System.Type.GetType("System.String"));//0
            dtActual.Columns.Add("Year", System.Type.GetType("System.Decimal"));//0
            dtActual.Columns.Add("Num of Week", System.Type.GetType("System.String"));//0
            dtActual.Columns.Add("TAC ID", System.Type.GetType("System.Decimal"));//0
            dtActual.Columns.Add("TAC Code", System.Type.GetType("System.String"));//0
            dtActual.Columns.Add("Used IMEI Quantity", System.Type.GetType("System.Decimal"));//0
            dtActual.Columns.Add("Init Number", System.Type.GetType("System.Decimal"));

            dtOrder.Columns.Add("ID", System.Type.GetType("System.Decimal"));//
            dtOrder.Columns.Add("Product ID", System.Type.GetType("System.Decimal"));
            dtOrder.Columns.Add("Product Code", System.Type.GetType("System.String"));
            dtOrder.Columns.Add("Customer ID", System.Type.GetType("System.Decimal"));
            dtOrder.Columns.Add("Customer Code", System.Type.GetType("System.String"));
            dtOrder.Columns.Add("Year", System.Type.GetType("System.Decimal"));
            dtOrder.Columns.Add("Num of Week", System.Type.GetType("System.String"));
            dtOrder.Columns.Add("Numbers", System.Type.GetType("System.Decimal"));
            dtOrder.Columns.Add("backlog", System.Type.GetType("System.Decimal"));

            dtBacklog.Columns.Add("ID", System.Type.GetType("System.Decimal"));//
            dtBacklog.Columns.Add("Product ID", System.Type.GetType("System.Decimal"));
            dtBacklog.Columns.Add("Product Code", System.Type.GetType("System.String"));
            dtBacklog.Columns.Add("Customer ID", System.Type.GetType("System.Decimal"));
            dtBacklog.Columns.Add("Customer Code", System.Type.GetType("System.String"));
            dtBacklog.Columns.Add("Year", System.Type.GetType("System.Decimal"));
            dtBacklog.Columns.Add("Num of Week", System.Type.GetType("System.String"));
            dtBacklog.Columns.Add("Numbers", System.Type.GetType("System.Decimal"));
            dtBacklog.Columns.Add("backlog", System.Type.GetType("System.Decimal"));

        }

        private void btnLoad_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialogOutput = new OpenFileDialog();
            openFileDialogOutput.Filter = "EXCEL files(*.xlsx)|*.xlsx|EXCEL files(*.xls)|*.xls";//
            openFileDialogOutput.FilterIndex = 0;
            openFileDialogOutput.RestoreDirectory = true;

            if (openFileDialogOutput.ShowDialog() != DialogResult.OK)
                return;



            btnLoad.Enabled = false;

            ExcelPathName = openFileDialogOutput.FileName;

            this.Text = "Import log:" + ExcelPathName;

            toolStripProgressBarALLValue = 0;
            toolStripProgressBarExcelValue = 0;
            toolStripProgressBarALL.Value = 0;
            toolStripProgressBarExcel.Value = 0;

            textBoxLOGText = "Import Data：\r\n filename:" + ExcelPathName + "\r\n\r\n";
            textBoxERRORText = "Error Note：\r\n filename:" + ExcelPathName + "\r\n\r\n";

            textBoxLOG.Text = "Import Data：\r\n filename:" + ExcelPathName + "\r\n\r\n";
            textBoxERROR.Text = "Error Note：\r\n filename:" + ExcelPathName + "\r\n\r\n";

            backgroundWorkerImport.RunWorkerAsync(this);
        }

        private void backgroundWorkerImport_DoWork(object sender, System.ComponentModel.DoWorkEventArgs e)
        {
            int i, j, iTemp = 0, iTemp1 = 0, iTemp2=0, iTemp3=0;
            bool bTemp = true;
            int iSheet, iCounter;
            DataSet dsCSV = new DataSet();
            int iI = 0, iP = 0, k;
            string sICode = "", sPCode = "";
            int iTAC = 0,iInitNumber=0;
            string sTAC = "";
            string sTemp = "";
            string sNowWeek = ClassIm.GetWeek(dateTimePickerP.Value).ToString();
            int iYear = 0, iWeek = 0;
            object[] oTemp1 = new object[7];
            object[] oTemp2 = new object[3];
            object[] oTemp3 = new object[7];
            object[] oTemp4 = new object[14];
            object[] oTemp5 = new object[14];
            object[] oTemp6 = new object[9];
            object objOpt = System.Reflection.Missing.Value;

            dtProduct.Clear(); dtCustumor.Clear(); dtAcquire.Clear(); dtActual.Clear(); dtOrder.Clear(); dtTAC.Clear();

            Microsoft.Office.Interop.Excel.Application m_xlsApp = null;
            Workbook m_Workbook = null;
            Worksheet m_Worksheet = null;

            //toolStripProgressBarALL.Value = 0;
            //toolStripProgressBarExcel.Value = 0;

            //textBoxLOG.Text = "Import Data：\r\n filename:" + ExcelPathName + "\r\n\r\n";
            //textBoxERROR.Text = "Error Note：\r\n filename:" + ExcelPathName + "\r\n\r\n";

           
            try
            {
                
                m_xlsApp = new Microsoft.Office.Interop.Excel.Application();
                m_Workbook = m_xlsApp.Workbooks.Open(ExcelPathName, objOpt, objOpt, objOpt, objOpt, objOpt, objOpt, objOpt, objOpt, objOpt, objOpt, objOpt, objOpt, objOpt, objOpt);

                //toolStripProgressBarALLMaximum = m_Workbook.Worksheets.Count + 2;
                toolStripProgressBarALLMaximum = 7 + 2;
                backgroundWorkerImport.ReportProgress(0);

                #region Index
                //定位表格
                iProduct = -1; iCustomer = -1; iTAC = -1; iAcquire = -1; iActual = -1; iSAPorder = -1; iSalesplan = -1;
                for (iSheet = 1; iSheet <= m_Workbook.Worksheets.Count; iSheet++)
                {
                    m_Worksheet = (Microsoft.Office.Interop.Excel.Worksheet)m_Workbook.Worksheets.get_Item(iSheet);

                    if (m_Worksheet.Name.ToUpper().IndexOf(sProduct.ToUpper()) >= 0)
                    {
                        iProduct = iSheet;
                    }

                    if (m_Worksheet.Name.ToUpper().IndexOf(sCustomer.ToUpper()) >= 0)
                    {
                        iCustomer = iSheet;
                    }

                    if (m_Worksheet.Name.ToUpper().IndexOf(sTACNumber.ToUpper()) >= 0)
                    {
                        iTACNumber = iSheet;
                    }

                    if (m_Worksheet.Name.ToUpper().IndexOf(sAcquire.ToUpper()) >= 0)
                    {
                        iAcquire = iSheet;
                    }

                    if (m_Worksheet.Name.ToUpper().IndexOf(sActual.ToUpper()) >= 0)
                    {
                        iActual = iSheet;
                    }

                    if (m_Worksheet.Name.ToUpper().IndexOf(sSAPorder.ToUpper()) >= 0)
                    {
                        iSAPorder = iSheet;
                    }

                    if (m_Worksheet.Name.ToUpper().IndexOf(sSalesplan.ToUpper()) >= 0)
                    {
                        iSalesplan = iSheet;
                    }


                }

                //定位表格列
                if (iProduct != -1)
                {
                    m_Worksheet = (Microsoft.Office.Interop.Excel.Worksheet)m_Workbook.Worksheets.get_Item(iProduct);
                    iTableProduct[0] = 2; iTableProduct[1] = 4; iTableProduct[2] = 8; iTableProduct[3] = 7; iTableProduct[4] = 6; iTableProduct[5] = 3;
                    for (i = 1; i <= m_Worksheet.UsedRange.Columns.Count; i++)
                    {
                        for (j = 0; j < sTableProduct.Length; j++)
                        {
                            if(((Range)(m_Worksheet.Cells[1, i])).Text.ToString().ToUpper()==sTableProduct[j].ToUpper())
                            {
                                iTableProduct[j] = i;
                            }
                        }
                    }

                }

                if (iCustomer != -1)
                {
                    m_Worksheet = (Microsoft.Office.Interop.Excel.Worksheet)m_Workbook.Worksheets.get_Item(iCustomer);
                    iTableCustomer[0] = 1; iTableCustomer[1] = 2;
                    for (i = 1; i <= m_Worksheet.UsedRange.Columns.Count; i++)
                    {
                        for (j = 0; j < sTableCustomer.Length; j++)
                        {
                            if (((Range)(m_Worksheet.Cells[1, i])).Text.ToString().ToUpper() == sTableCustomer[j].ToUpper())
                            {
                                iTableCustomer[j] = i;
                            }
                        }
                    }

                }

                if (iTACNumber != -1)
                {
                    m_Worksheet = (Microsoft.Office.Interop.Excel.Worksheet)m_Workbook.Worksheets.get_Item(iTACNumber);
                    iTableTAC[0] = 9; iTableTAC[1] = 2; iTableTAC[2] = 5;
                    for (i = 1; i <= m_Worksheet.UsedRange.Columns.Count; i++)
                    {
                        for (j = 0; j < sTableTAC.Length; j++)
                        {
                            if (((Range)(m_Worksheet.Cells[1, i])).Text.ToString().ToUpper() == sTableTAC[j].ToUpper())
                            {
                                iTableTAC[j] = i;
                            }
                        }
                    }

                }

                if (iAcquire != -1)
                {
                    m_Worksheet = (Microsoft.Office.Interop.Excel.Worksheet)m_Workbook.Worksheets.get_Item(iAcquire);
                    iTableAcquire[0] = 9; iTableAcquire[1] = 2; iTableAcquire[2] = 5; iTableAcquire[3] = 10;
                    for (i = 1; i <= m_Worksheet.UsedRange.Columns.Count; i++)
                    {
                        for (j = 0; j < sTableAcquire.Length; j++)
                        {
                            if (((Range)(m_Worksheet.Cells[1, i])).Text.ToString().ToUpper() == sTableAcquire[j].ToUpper())
                            {
                                iTableAcquire[j] = i;
                            }
                        }
                    }

                    //取得申请区间index
                    listIndex.Clear();
                    iTemp3 = 0; iTemp = 1;
                    while (true)
                    {
                        iTemp3++;
                        iTemp1 = 0;

                        for (k = 1; k <= m_Worksheet.UsedRange.Columns.Count; k++)
                        {
                            if (((Range)(m_Worksheet.Cells[1, k])).Text.ToString().ToUpper() == (sTableAcquire[3].ToUpper() + iTemp.ToString() + " From").ToUpper())
                            {
                                iTemp1 = k;
                                break;
                            }
                        }

                        iTemp2 = 0;
                        for (k = 1; k <= m_Worksheet.UsedRange.Columns.Count; k++)
                        {
                            if (((Range)(m_Worksheet.Cells[1, k])).Text.ToString().ToUpper() == (sTableAcquire[3].ToUpper() + iTemp.ToString() + " To").ToUpper())
                            {
                                iTemp2 = k;
                                break;
                            }
                        }

                        if (iTemp1 * iTemp2 == 0) //已没有申请
                            break;

                        ClassIndexofOrder ciofOrder = new ClassIndexofOrder(iTemp1, iTemp2, iTemp3);
                        listIndex.Add(ciofOrder);

                        iTemp ++;
                    }


                }

                if (iActual != -1)
                {
                    m_Worksheet = (Microsoft.Office.Interop.Excel.Worksheet)m_Workbook.Worksheets.get_Item(iActual);
                    iTableActual[0] = 1; iTableActual[1] = 3; iTableActual[2] = 8; iTableActual[3] = 14;
                    for (i = 1; i <= m_Worksheet.UsedRange.Columns.Count; i++)
                    {
                        for (j = 0; j < sTableActual.Length; j++)
                        {
                            if (((Range)(m_Worksheet.Cells[1, i])).Text.ToString().ToUpper() == sTableActual[j].ToUpper())
                            {
                                iTableActual[j] = i;
                            }
                        }
                    }

                }

                if (iSAPorder != -1)
                {
                    m_Worksheet = (Microsoft.Office.Interop.Excel.Worksheet)m_Workbook.Worksheets.get_Item(iSAPorder);
                    iTableSAPorder[0] = 9; iTableSAPorder[1] = 5; iTableSAPorder[2] = 11; iTableSAPorder[3] = 14;
                    for (i = 1; i <= m_Worksheet.UsedRange.Columns.Count; i++)
                    {
                        for (j = 0; j < sTableSAPorder.Length; j++)
                        {
                            if (((Range)(m_Worksheet.Cells[1, i])).Text.ToString().ToUpper() == sTableSAPorder[j].ToUpper())
                            {
                                iTableSAPorder[j] = i;
                            }
                        }
                    }

                }

                if (iSalesplan != -1)
                {
                    m_Worksheet = (Microsoft.Office.Interop.Excel.Worksheet)m_Workbook.Worksheets.get_Item(iSalesplan);
                    iTableSalesplan[0] = 2; iTableSalesplan[1] = 1;
                    for (i = 1; i <= m_Worksheet.UsedRange.Columns.Count; i++)
                    {
                        for (j = 0; j < sTableSalesplan.Length; j++)
                        {
                            if (((Range)(m_Worksheet.Cells[1, i])).Text.ToString().ToUpper() == sTableSalesplan[j].ToUpper())
                            {
                                iTableSalesplan[j] = i;
                            }
                        }
                    }

                }




                #endregion

                #region product
                dtProduct.Clear();
                if (iProduct > 0)
                {
                    m_Worksheet = (Microsoft.Office.Interop.Excel.Worksheet)m_Workbook.Worksheets.get_Item(iProduct);
                    toolStripProgressBarALLValue++;
                    textBoxLOGText += "\r\nRead TABLE '" + m_Worksheet.Name + "'\r\n";

                    toolStripProgressBarExcelMaximum = m_Worksheet.UsedRange.Rows.Count + 1;
                    toolStripProgressBarExcelValue = 0;
                    backgroundWorkerImport.ReportProgress(0);

                    for (i = 2; i <= m_Worksheet.UsedRange.Rows.Count; i++)
                    {
                        toolStripProgressBarExcelValue++;
                        toolStripStatusLabelWarnText = "LINE:" + i.ToString();
                        backgroundWorkerImport.ReportProgress(0);

                        //跳过空行
                        if (((Range)(m_Worksheet.Cells[i, iTableProduct[0]])) == null)
                        {
                            continue;
                        }

                        if (((Range)(m_Worksheet.Cells[i, iTableProduct[0]])).Text.ToString() == "")
                            continue;

                        //隐藏行
                        if((bool)m_Worksheet.get_Range(m_Worksheet.Rows[i, objOpt], m_Worksheet.Rows[i, objOpt]).Hidden && !checkBoxHidden.Checked)
                            continue;

                        //检查
                        var qProduct = from dt in dtProduct.AsEnumerable()//
                                       where (dt.Field<string>("Product Code") == ((Range)(m_Worksheet.Cells[i, iTableProduct[0]])).Text.ToString())
                                       select dt;

                        if (qProduct.Count() <= 0)
                        {
                            oTemp1[0] = i;
                            oTemp1[1] = ((Range)(m_Worksheet.Cells[i, iTableProduct[0]])).Text.ToString();//code
                            oTemp1[2] = ((Range)(m_Worksheet.Cells[i, iTableProduct[1]])).Text.ToString().Trim();//name
                            if (((Range)(m_Worksheet.Cells[i, iTableProduct[2]])).Text.ToString().ToUpper() == "Y")
                                oTemp1[3] = 2;
                            else
                                oTemp1[3] = 1;

                            oTemp1[4] = 0;
                            //sTemp = ((Range)(m_Worksheet.Cells[i, 3])).Text.ToString();

                            if (((Range)(m_Worksheet.Cells[i, iTableProduct[3]])).Text.ToString() != "")
                                oTemp1[4] = decimal.Parse(((Range)(m_Worksheet.Cells[i, iTableProduct[3]])).Text.ToString());
                            oTemp1[5] = ((Range)(m_Worksheet.Cells[i, iTableProduct[4]])).Text.ToString();
                            oTemp1[6] = ((Range)(m_Worksheet.Cells[i, iTableProduct[5]])).Text.ToString();

                            dtProduct.Rows.Add(oTemp1);
                        }
                    }
                }
                #endregion
                
                #region Customer
                dtCustumor.Clear();
                if (iCustomer > 0)
                {
                    m_Worksheet = (Microsoft.Office.Interop.Excel.Worksheet)m_Workbook.Worksheets.get_Item(iCustomer);

                    toolStripProgressBarALLValue++;
                    textBoxLOGText += "\r\nRead TABLE '" + m_Worksheet.Name + "'\r\n";
                    backgroundWorkerImport.ReportProgress(0);

                    toolStripProgressBarExcelMaximum = m_Worksheet.UsedRange.Rows.Count + 1;
                    toolStripProgressBarExcelValue = 0;
                    backgroundWorkerImport.ReportProgress(0);

                    for (i = 2; i <= m_Worksheet.UsedRange.Rows.Count; i++)
                    {
                        toolStripProgressBarExcelValue++;
                        toolStripStatusLabelWarnText = "LINE:" + i.ToString();
                        backgroundWorkerImport.ReportProgress(0);

                        //跳过空行
                        if (((Range)(m_Worksheet.Cells[i, iTableCustomer[1]])) == null)
                        {
                            continue;
                        }

                        if (((Range)(m_Worksheet.Cells[i, iTableCustomer[1]])).Text.ToString() == "" || ((Range)(m_Worksheet.Cells[i, iTableCustomer[0]])).Text.ToString() == "")
                            continue;


                        //隐藏行
                        if ((bool)m_Worksheet.get_Range(m_Worksheet.Rows[i, objOpt], m_Worksheet.Rows[i, objOpt]).Hidden && !checkBoxHidden.Checked)
                            continue;


                        //检查
                        var qCustomer = from dt in dtCustumor.AsEnumerable()//
                                        where (dt.Field<string>("Customer Code") == ((Range)(m_Worksheet.Cells[i, iTableCustomer[1]])).Text.ToString())
                                        select dt;

                        if (qCustomer.Count() <= 0)
                        {
                            oTemp2[0] = i;
                            oTemp2[1] = ((Range)(m_Worksheet.Cells[i, iTableCustomer[1]])).Text.ToString();
                            oTemp2[2] = ((Range)(m_Worksheet.Cells[i, iTableCustomer[0]])).Text.ToString();

                            dtCustumor.Rows.Add(oTemp2);
                        }
                    }
                    
                }

                #endregion

                #region TAC
                dtTAC.Clear();
                if (iTACNumber > 0)
                {
                    m_Worksheet = (Microsoft.Office.Interop.Excel.Worksheet)m_Workbook.Worksheets.get_Item(iTACNumber);
                    toolStripProgressBarALLValue++;
                    textBoxLOGText += "\r\nRead TABLE for TAC '" + m_Worksheet.Name + "'\r\n";
                    backgroundWorkerImport.ReportProgress(0);

                    toolStripProgressBarExcelMaximum = m_Worksheet.UsedRange.Rows.Count + 1;
                    toolStripProgressBarExcelValue = 0;
                    backgroundWorkerImport.ReportProgress(0);


                    for (i = 2; i <= m_Worksheet.UsedRange.Rows.Count; i++)
                    {
                        toolStripProgressBarExcelValue = i;
                        toolStripStatusLabelWarnText = "LINE:" + i.ToString();
                        backgroundWorkerImport.ReportProgress(0);
                        //跳过空行
                        if (((Range)(m_Worksheet.Cells[i, iTableTAC[0]])) == null)
                        {
                            continue;
                        }

                        if (((Range)(m_Worksheet.Cells[i, iTableTAC[0]])).Text.ToString() == "")
                            continue;


                        //隐藏行
                        if ((bool)m_Worksheet.get_Range(m_Worksheet.Rows[i, objOpt], m_Worksheet.Rows[i, objOpt]).Hidden && !checkBoxHidden.Checked)
                            continue;

                        //iICODE
                        iI = 0; sICode = "";
                        var qCustomer = from dt in dtCustumor.AsEnumerable()//
                                        where (dt.Field<string>("Customer Name").ToUpper() == ((Range)(m_Worksheet.Cells[i, iTableTAC[2]])).Text.ToString().ToUpper().Trim())
                                        select dt;
                        if (qCustomer.Count() <= 0)
                        {
                            string sss = ((Range)(m_Worksheet.Cells[i, iTableTAC[2]])).Text.ToString().ToUpper();
                            textBoxERRORText += m_Worksheet.Name + " Line " + i.ToString() + ": NO Customers Code\r\n";
                            backgroundWorkerImport.ReportProgress(0);
                            continue;
                        }
                        else
                        {
                            foreach (var item in qCustomer)//显示查询结果
                            {
                                iI = int.Parse(item.Field<decimal>("ID").ToString());
                                sICode = item.Field<string>("Customer Code");
                                break;
                            }
                        }


                        //iPCode
                        iP = 0; sPCode = "";
                        var qProduct = from dt in dtProduct.AsEnumerable()//
                                       where (dt.Field<string>("Product Code") == ClassIm.GetPCode(((Range)(m_Worksheet.Cells[i, iTableTAC[1]])).Text.ToString()).Trim())
                                       select dt;
                        if (qProduct.Count() <= 0)
                        {
                            textBoxERRORText += m_Worksheet.Name + " Line " + i.ToString() + ": NO Product Code\r\n";
                            backgroundWorkerImport.ReportProgress(0);
                            continue;
                        }
                        else
                        {
                            foreach (var item in qProduct)//显示查询结果
                            {
                                iP = int.Parse(item.Field<decimal>("ID").ToString());
                                sPCode = item.Field<string>("Product Code");
                                break;
                            }
                        }

                        //TAC
                        iTAC = 0; sTAC = "";
                        var qTAC = from dt in dtTAC.AsEnumerable()//
                                   where (dt.Field<string>("Product Code") == sPCode) && (dt.Field<string>("Customer Code") == sICode) && (dt.Field<string>("TAC Code") == ((Range)(m_Worksheet.Cells[i, iTableTAC[0]])).Text.ToString())
                                   select dt;

                        if (qTAC.Count() <= 0)
                        {
                            var qTAC1 = from dt in dtTAC.AsEnumerable()//
                                       where (dt.Field<string>("Product Code") == sPCode) && (dt.Field<string>("Customer Code") == sICode)
                                       orderby dt.Field<decimal>("Init Number") descending 
                                       select dt;

                            if (qTAC1.Count() <= 0) //没有以前的记录
                            {
                                oTemp3[0] = i;
                                oTemp3[1] = ((Range)(m_Worksheet.Cells[i, iTableTAC[0]])).Text.ToString();
                                oTemp3[2] = iP;
                                oTemp3[3] = sPCode;
                                oTemp3[4] = iI;
                                oTemp3[5] = sICode;
                                oTemp3[6] = 0;
                            }
                            else
                            {
                                oTemp3[0] = i;
                                oTemp3[1] = ((Range)(m_Worksheet.Cells[i, iTableTAC[0]])).Text.ToString();
                                oTemp3[2] = iP;
                                oTemp3[3] = sPCode;
                                oTemp3[4] = iI;
                                oTemp3[5] = sICode;
                                foreach (var item in qTAC1)//查询结果
                                {
                                    oTemp3[6] = decimal.Parse(item.Field<decimal>("Init Number").ToString()) + (decimal)iRange;
                                    break;
                                }
                            }

                            dtTAC.Rows.Add(oTemp3);
                        }


                    }


                }
                #endregion

                #region accquire
                dtAcquire.Clear();
                if (iAcquire > 0)
                {
                    m_Worksheet = (Microsoft.Office.Interop.Excel.Worksheet)m_Workbook.Worksheets.get_Item(iAcquire);
                    toolStripProgressBarALLValue++;
                    textBoxLOGText += "\r\nRead TABLE for Request '" + m_Worksheet.Name + "'\r\n";
                    backgroundWorkerImport.ReportProgress(0);

                    toolStripProgressBarExcelMaximum = m_Worksheet.UsedRange.Rows.Count + 1;
                    toolStripProgressBarExcelValue = 0;
                    backgroundWorkerImport.ReportProgress(0);

                    iCounter = 1;
                    for (i = 2; i <= m_Worksheet.UsedRange.Rows.Count; i++)
                    {
                        toolStripProgressBarExcelValue = i;
                        toolStripStatusLabelWarnText = "LINE:" + i.ToString();
                        backgroundWorkerImport.ReportProgress(0);
                        //跳过空行
                        if (((Range)(m_Worksheet.Cells[i, iTableAcquire[0]])).Value2 == null || ((Range)(m_Worksheet.Cells[i, iTableAcquire[1]])).Value2 == null || ((Range)(m_Worksheet.Cells[i, iTableAcquire[2]])) == null)
                        {
                            continue;
                        }

                        if (((Range)(m_Worksheet.Cells[i, iTableAcquire[0]])).Text.ToString() == "" || ((Range)(m_Worksheet.Cells[i, iTableAcquire[1]])).Text.ToString() == "" || ((Range)(m_Worksheet.Cells[i, iTableAcquire[2]])).Text.ToString() == "")
                            continue;


                        //隐藏行
                        if ((bool)m_Worksheet.get_Range(m_Worksheet.Rows[i, objOpt], m_Worksheet.Rows[i, objOpt]).Hidden && !checkBoxHidden.Checked)
                            continue;

                        //iICODE
                        iI = 0; sICode = "";
                        var qCustomer = from dt in dtCustumor.AsEnumerable()//
                                        where (dt.Field<string>("Customer Name").ToUpper() == ((Range)(m_Worksheet.Cells[i, iTableAcquire[2]])).Text.ToString().ToUpper().Trim())
                                        select dt;
                        if (qCustomer.Count() <= 0)
                        {
                            textBoxERRORText += m_Worksheet.Name + " Line " + i.ToString() + ": NO Customers Code\r\n";
                            backgroundWorkerImport.ReportProgress(0);
                            continue;
                        }
                        else
                        {
                            foreach (var item in qCustomer)//显示查询结果
                            {
                                iI = int.Parse(item.Field<decimal>("ID").ToString());
                                sICode = item.Field<string>("Customer Code");
                                break;
                            }
                        }


                        //iPCode
                        iP = 0; sPCode = "";
                        var qProduct = from dt in dtProduct.AsEnumerable()//
                                       where (dt.Field<string>("Product Code") == ClassIm.GetPCode(((Range)(m_Worksheet.Cells[i, iTableAcquire[1]])).Text.ToString().Trim()))
                                       select dt;
                        if (qProduct.Count() <= 0)
                        {
                            textBoxERRORText += m_Worksheet.Name + " Line " + i.ToString() + ": NO Product Code\r\n";
                            backgroundWorkerImport.ReportProgress(0);
                            continue;
                        }
                        else
                        {
                            foreach (var item in qProduct)//显示查询结果
                            {
                                iP = int.Parse(item.Field<decimal>("ID").ToString());
                                sPCode = item.Field<string>("Product Code");
                                break;
                            }
                        }

                        //检查TAC
                        iTAC = 0; sTAC = ""; iInitNumber = 0;
                        var qTAC = from dt2 in dtTAC.AsEnumerable()//得到TAC初始值
                                   where (dt2.Field<string>("Product Code") == sPCode) && (dt2.Field<string>("Customer Code") == sICode) && (dt2.Field<string>("TAC Code").ToUpper()==((Range)(m_Worksheet.Cells[i, iTableAcquire[0]])).Text.ToString().ToUpper().Trim())//条件
                                   select dt2;
                        if (qTAC.Count() <= 0)
                        {
                            continue;
                        }
                        foreach (var itemTAC in qTAC)//显示查询结果//有TAC
                        {
                            iTAC = int.Parse(itemTAC.Field<decimal>("ID").ToString());
                            sTAC = itemTAC.Field<string>("TAC Code");
                            iInitNumber = (int)itemTAC.Field<decimal>("Init Number");
                            break;
                        }

                        //iTemp = 1;
                        //while (true)
                        //{
                        //    iTemp1 = 0;

                        //    for (k = 1; k <= m_Worksheet.UsedRange.Columns.Count; k++)
                        //    {
                        //        if (((Range)(m_Worksheet.Cells[1, k])).Text.ToString().ToUpper() == (sTableAcquire[3].ToUpper() + iTemp.ToString() + " From").ToUpper())
                        //        {
                        //            iTemp1 = k;
                        //            break;
                        //        }
                        //    }

                        //    iTemp2 = 0;
                        //    for (k = 1; k <= m_Worksheet.UsedRange.Columns.Count; k++)
                        //    {
                        //        if (((Range)(m_Worksheet.Cells[1, k])).Text.ToString().ToUpper() == (sTableAcquire[3].ToUpper() + iTemp.ToString() + " To").ToUpper())
                        //        {
                        //            iTemp2 = k;
                        //            break;
                        //        }
                        //    }

                        //    if (iTemp1 * iTemp2 == 0) //已没有申请
                        //        break;

                        //    iTemp++;
                            //判断合理性
                            //跳过空行
                        foreach(ClassIndexofOrder ciofOrder in listIndex)
                        {
                            iTemp1 = ciofOrder.iFrom;
                            iTemp2 = ciofOrder.iTo;

                            if (((Range)(m_Worksheet.Cells[i, iTemp1])).Value2 == null || ((Range)(m_Worksheet.Cells[i, iTemp2])).Value2 == null)
                            {
                                continue;
                            }
                            if (((Range)(m_Worksheet.Cells[i, iTemp1])).Value2.ToString() == "" || ((Range)(m_Worksheet.Cells[i, iTemp2])).Value2.ToString() == "")
                            {
                                continue;
                            }

                            if (decimal.Parse(((Range)(m_Worksheet.Cells[i, iTemp2])).Value2.ToString()) <= decimal.Parse(((Range)(m_Worksheet.Cells[i, iTemp1])).Value2.ToString()))//区间不正确
                            {
                                textBoxERRORText += m_Worksheet.Name + " Line " + i.ToString() + ": Data Error - request data wrong\r\n";
                                backgroundWorkerImport.ReportProgress(0);
                                continue;
                            }

                            //取得最后申请数
                            var qAcquire = from dt in dtAcquire.AsEnumerable()//
                                           where (dt.Field<string>("Product Code") == sPCode) && (dt.Field<string>("Customer Code") == sICode) && (dt.Field<string>("TAC Code") == sTAC)
                                           orderby dt.Field<decimal>("End Number") descending 
                                           select dt;
                            if (qAcquire.Count() <= 0) //第一个区间
                            {
                                oTemp4[0] = iCounter;
                                oTemp4[1] = iP;
                                oTemp4[2] = sPCode;
                                oTemp4[3] = iI;
                                oTemp4[4] = sICode;
                                oTemp4[5] = int.Parse(((Range)(m_Worksheet.Cells[i, iTemp2])).Value2.ToString());
                                oTemp4[6] = dateTimePickerP.Value.Year;
                                oTemp4[7] = sNowWeek;
                                oTemp4[8] = int.Parse(((Range)(m_Worksheet.Cells[i, iTemp1])).Value2.ToString());
                                oTemp4[9] = dateTimePickerP.Value.ToShortDateString();
                                oTemp4[10] = iTAC;
                                oTemp4[11] = sTAC;

                                oTemp4[12] = iInitNumber;
                                oTemp4[13] = 0;

                                dtAcquire.Rows.Add(oTemp4);
                                iCounter++;
                            }
                            else //不是第一区间，检查合理性
                            {
                                foreach (var item in qAcquire)//显示查询结果
                                {

                                    if (decimal.Parse(((Range)(m_Worksheet.Cells[i, iTemp1])).Value2.ToString()) <= decimal.Parse(item.Field<decimal>("End Number").ToString()))//区间不正确
                                    {
                                        textBoxERRORText += m_Worksheet.Name + " Line " + i.ToString() + ": Data Error - request data wrong\r\n";
                                        backgroundWorkerImport.ReportProgress(0);
                                        break;
                                    }
                                    oTemp4[0] = iCounter;
                                    oTemp4[1] = iP;
                                    oTemp4[2] = sPCode;
                                    oTemp4[3] = iI;
                                    oTemp4[4] = sICode;
                                    oTemp4[5] = int.Parse(((Range)(m_Worksheet.Cells[i, iTemp2])).Value2.ToString());
                                    oTemp4[6] = dateTimePickerP.Value.Year;
                                    oTemp4[7] = sNowWeek;
                                    oTemp4[8] = int.Parse(((Range)(m_Worksheet.Cells[i, iTemp1])).Value2.ToString());
                                    oTemp4[9] = dateTimePickerP.Value.ToShortDateString();
                                    oTemp4[10] = iTAC;
                                    oTemp4[11] = sTAC;
                                    oTemp4[12] = iInitNumber;
                                    oTemp4[13] = 0;

                                    dtAcquire.Rows.Add(oTemp4);
                                    iCounter++;

                                    break;
                                }
                            }

                        }



                    }


                }
                #endregion

                #region manufacture & backlog
                dtActual.Clear();
                dtBacklog.Clear();
                if (iActual > 0)
                {
                    m_Worksheet = (Microsoft.Office.Interop.Excel.Worksheet)m_Workbook.Worksheets.get_Item(iActual);
                    toolStripProgressBarALLValue++;
                    textBoxLOGText += "\r\nRead TABLE for CW & backlog '" + m_Worksheet.Name + "'\r\n";
                    backgroundWorkerImport.ReportProgress(0);

                    toolStripProgressBarExcelMaximum = m_Worksheet.UsedRange.Rows.Count + 1;
                    toolStripProgressBarExcelValue = 0;
                    backgroundWorkerImport.ReportProgress(0);
                    for (i = 2; i <= m_Worksheet.UsedRange.Rows.Count; i++)
                    {
                        toolStripProgressBarExcelValue = i;
                        toolStripStatusLabelWarnText = "LINE:" + i.ToString();
                        backgroundWorkerImport.ReportProgress(0);
                        //跳过空行
                        if (((Range)(m_Worksheet.Cells[i, iTableActual[0]])).Value2 == null || ((Range)(m_Worksheet.Cells[i, iTableActual[1]])).Value2 == null)
                        {
                            continue;
                        }

                        if (((Range)(m_Worksheet.Cells[i, iTableActual[0]])).Text.ToString() == "" || ((Range)(m_Worksheet.Cells[i, iTableActual[1]])).Text.ToString() == "")
                            continue;


                        //隐藏行
                        if ((bool)m_Worksheet.get_Range(m_Worksheet.Rows[i, objOpt], m_Worksheet.Rows[i, objOpt]).Hidden && !checkBoxHidden.Checked)
                            continue;

                        //if (((Range)(m_Worksheet.Cells[i, iTableActual[2]])).Value2.ToString().Length != 15)
                        if (((Range)(m_Worksheet.Cells[i, iTableActual[2]])).Value2.ToString().Length < 14)
                        {
                            //textBoxERRORText += m_Worksheet.Name + " Line " + i.ToString() + ": Data Error - Length is not 15\r\n";
                            textBoxERRORText += m_Worksheet.Name + " Line " + i.ToString() + ": Data Error - Length is less than 14\r\n";
                            backgroundWorkerImport.ReportProgress(0);
                            continue;
                        }

                        //iICODE customer use name
                        iI = 0; sICode = "";
                        
                        var qCustomer = from dt in dtCustumor.AsEnumerable()//
                                        where (dt.Field<string>("Customer Name").ToUpper() == ((Range)(m_Worksheet.Cells[i, iTableActual[1]])).Text.ToString().ToUpper().Trim())
                                        select dt;
                        if (qCustomer.Count() <= 0)
                        {
                            textBoxERRORText += m_Worksheet.Name + " Line " + i.ToString() + ": NO Customers Code\r\n";
                            backgroundWorkerImport.ReportProgress(0);
                            continue;
                        }
                        else
                        {
                            foreach (var item in qCustomer)//显示查询结果
                            {
                                iI = int.Parse(item.Field<decimal>("ID").ToString());
                                sICode = item.Field<string>("Customer Code");
                                break;
                            }
                        }


                        //iPCode product use Type Designation
                        iP = 0; sPCode = "";
                        var qProduct = from dt in dtProduct.AsEnumerable()//
                                       where (dt.Field<string>("Type Designation").ToUpper() == ClassIm.GetPCode(((Range)(m_Worksheet.Cells[i, iTableActual[0]])).Text.ToString().ToUpper().Trim()))
                                       select dt;
                        if (qProduct.Count() <= 0)
                        {
                            textBoxERRORText += m_Worksheet.Name + " Line " + i.ToString() + ": NO Product Code\r\n";
                            backgroundWorkerImport.ReportProgress(0);
                            continue;
                        }
                        else
                        {
                            foreach (var item in qProduct)//显示查询结果
                            {
                                iP = int.Parse(item.Field<decimal>("ID").ToString());
                                sPCode = item.Field<string>("Product Code");
                                break;
                            }
                        }

                        //manufacture
                        iTemp1 = int.Parse(((Range)(m_Worksheet.Cells[i, iTableActual[2]])).Value2.ToString().Substring(8, 6));
                        sTAC = ((Range)(m_Worksheet.Cells[i, iTableActual[2]])).Value2.ToString().Substring(0, 8);

                        //检查TAC
                        iTAC = 0;  iInitNumber = 0;
                        var qTAC = from dt2 in dtTAC.AsEnumerable()//得到TAC初始值
                                   where (dt2.Field<string>("Product Code") == sPCode) && (dt2.Field<string>("Customer Code") == sICode) && (dt2.Field<string>("TAC Code").ToUpper() == sTAC.ToUpper())//条件
                                   select dt2;
                        if (qTAC.Count() <= 0)
                        {
                            textBoxERRORText += m_Worksheet.Name + " Line " + i.ToString() + ": NO TAC Code\r\n";
                            backgroundWorkerImport.ReportProgress(0);
                            continue;
                        }
                        foreach (var itemTAC in qTAC)//显示查询结果//有TAC
                        {
                            iTAC = int.Parse(itemTAC.Field<decimal>("ID").ToString());
                            sTAC = itemTAC.Field<string>("TAC Code");
                            iInitNumber = (int)itemTAC.Field<decimal>("Init Number");
                            break;
                        }

                        //var qActual = from dt2 in dtActual.AsEnumerable()//
                        //              where (dt2.Field<string>("Product Code") == sPCode) && (dt2.Field<string>("Customer Code") == sICode) && (dt2.Field<string>("TAC Code") == sTAC)//条件
                        //              select dt2;
                        var qActual = from dt2 in dtActual.AsEnumerable()//
                                      where (dt2.Field<string>("Product Code") == sPCode) && (dt2.Field<string>("Customer Code") == sICode)//条件
                                      select dt2;
                        if (qActual.Count() <= 0)
                        {
                            oTemp5[0] = iP;
                            oTemp5[1] = sPCode;
                            oTemp5[2] = iI;
                            oTemp5[3] = sICode;
                            oTemp5[4] = 0;
                            oTemp5[5] = iTemp1;
                            oTemp5[6] = i;
                            oTemp5[7] = dateTimePickerP.Value.ToShortDateString();
                            oTemp5[8] = dateTimePickerP.Value.Year;
                            oTemp5[9] = sNowWeek;
                            oTemp5[10] = iTAC;
                            oTemp5[11] = sTAC;
                            oTemp5[12] = iTemp1;
                            oTemp5[13] = iInitNumber;

                            dtActual.Rows.Add(oTemp5);
                        }
                        else //修改生产真实值,保证只有一个真实生产号
                        {
                            bTemp = true;
                            foreach(var resActual in qActual)                
                            {
                                iTemp=int.Parse(resActual.Field<decimal>("Init Number").ToString())+int.Parse(resActual.Field<decimal>("Init Number").ToString());
                                if (iInitNumber + iTemp1 > iTemp) //更大的生产号
                                {
                                    bTemp = false;

                                    //修改
                                    resActual.SetField<decimal>("End Number",iTemp1);
                                    resActual.SetField<decimal>("Acquire ID", i);
                                    resActual.SetField<decimal>("TAC ID", iTAC);
                                    resActual.SetField<string>("TAC Code", sTAC);
                                    resActual.SetField<decimal>("Used IMEI Quantity", iTemp1);
                                    resActual.SetField<decimal>("Init Number", iInitNumber);
                                }
                                else
                                {
                                }
                                break;
                            }
                            if (bTemp) //更小的生产号，跳过
                                continue;

                            if (((Range)(m_Worksheet.Cells[i, iTableActual[3]])).Text.ToString().Trim() == "")
                                iTemp3 = 0;
                            else
                            {
                                iTemp3 = int.Parse(((Range)(m_Worksheet.Cells[i, iTableActual[3]])).Value2.ToString().Trim());
                            }

                            //修改backlog，然后跳过
                            var qBacklog1 = from dt2 in dtBacklog.AsEnumerable()//
                                           where (dt2.Field<string>("Product Code") == sPCode) && (dt2.Field<string>("Customer Code") == sICode) && (dt2.Field<decimal>("Year") == dateTimePickerP.Value.Year) && (dt2.Field<string>("Num of Week") == sNowWeek)//条件
                                           select dt2;

                            foreach (var resBacklog1 in qBacklog1)
                            {
                                //修改
                                resBacklog1.SetField<decimal>("backlog", iTemp3);
                                break;
                            }

                            continue;

                        }


                        //backlog
                        if (((Range)(m_Worksheet.Cells[i, iTableActual[3]])).Text.ToString().Trim() == "")
                            iTemp1 = 0;
                        else
                        {
                            iTemp1 = int.Parse(((Range)(m_Worksheet.Cells[i, iTableActual[3]])).Value2.ToString().Trim());
                        }

                        var qBacklog = from dt2 in dtBacklog.AsEnumerable()//
                                       where (dt2.Field<string>("Product Code") == sPCode) && (dt2.Field<string>("Customer Code") == sICode) && (dt2.Field<decimal>("Year") == dateTimePickerP.Value.Year) && (dt2.Field<string>("Num of Week") == sNowWeek)//条件
                                       select dt2;
                        if (qBacklog.Count() <= 0)
                        {
                            oTemp6[0] = i;
                            oTemp6[1] = iP;
                            oTemp6[2] = sPCode;
                            oTemp6[3] = iI;
                            oTemp6[4] = sICode;
                            oTemp6[5] = dateTimePickerP.Value.Year;
                            oTemp6[6] = sNowWeek;
                            oTemp6[7] = 0;
                            oTemp6[8] = iTemp1;

                            dtBacklog.Rows.Add(oTemp6);
                        }





                    }
                }

                #endregion

                #region Order
                alSales.Clear();
                alSAP.Clear();
                dtOrder.Clear();

                #region SAP
                if (iSAPorder > 0)
                {
                    m_Worksheet = (Microsoft.Office.Interop.Excel.Worksheet)m_Workbook.Worksheets.get_Item(iSAPorder);
                    toolStripProgressBarALLValue++;
                    textBoxLOGText += "\r\nRead TABLE '" + m_Worksheet.Name + "'\r\n";
                    backgroundWorkerImport.ReportProgress(0);

                    toolStripProgressBarExcelMaximum = m_Worksheet.UsedRange.Rows.Count + 1;
                    toolStripProgressBarExcelValue = 0;
                    backgroundWorkerImport.ReportProgress(0);
                    for (i = 2; i <= m_Worksheet.UsedRange.Rows.Count; i++)
                    {
                        toolStripProgressBarExcelValue++;
                        toolStripStatusLabelWarnText = "LINE:" + i.ToString();
                        backgroundWorkerImport.ReportProgress(0);
                        
                        //跳过空行
                        if (((Range)(m_Worksheet.Cells[i, iTableSAPorder[3]])).Value2 == null)
                        {
                            break;
                        }

                        if (((Range)(m_Worksheet.Cells[i, iTableSAPorder[3]])).Value2.ToString() == "")
                        {
                            break;
                        }

                        //隐藏行
                        if ((bool)m_Worksheet.get_Range(m_Worksheet.Rows[i, objOpt], m_Worksheet.Rows[i, objOpt]).Hidden && !checkBoxHidden.Checked)
                            continue;

                        sPCode = ClassIm.GetPCode(((Range)(m_Worksheet.Cells[i, iTableSAPorder[0]])).Value2.ToString());
                        sICode = ((Range)(m_Worksheet.Cells[i, iTableSAPorder[1]])).Value2.ToString();
                        string[] sT = ((Range)(m_Worksheet.Cells[i, iTableSAPorder[2]])).Value2.ToString().Split('.');

                        //年为四位，判断哪个是年
                        if (sT[0].Length != 4)
                        {
                            if (sT[1].Length == 4 && sT[0].Length == 2)
                            {
                                sTemp = sT[0];
                                sT[0] = sT[1];
                                sT[1] = sTemp;
                            }
                            else
                            {
                                textBoxERRORText += m_Worksheet.Name + " Line " + i.ToString() + ": date format wrong '" + ((Range)(m_Worksheet.Cells[i, iTableSAPorder[2]])).Value2.ToString() + "'\r\n";
                                backgroundWorkerImport.ReportProgress(0);
                                continue;
                            }
                        }

                        var alQuery = from cImeOrder cime in alSAP
                                      where (cime.sPCode == sPCode && (cime.sICode == sICode) && (cime.iYear == int.Parse(sT[0])) && (cime.iWeek == int.Parse(sT[1])))
                                      select cime;

                        if (alQuery.Count() <= 0) //没有记录
                        {
                            //iICODE
                            iI = 0;
                            var qCustomer = from dt in dtCustumor.AsEnumerable()//
                                            where (dt.Field<string>("Customer Code") == ((Range)(m_Worksheet.Cells[i, iTableSAPorder[1]])).Text.ToString())
                                            select dt;
                            if (qCustomer.Count() <= 0)
                            {
                                textBoxERRORText += m_Worksheet.Name + " Line " + i.ToString() + ": NO Customers Code '" + ((Range)(m_Worksheet.Cells[i, iTableSAPorder[1]])).Text.ToString() + "'\r\n";
                                backgroundWorkerImport.ReportProgress(0);
                                continue;
                            }
                            else
                            {
                                foreach (var item in qCustomer)//显示查询结果
                                {
                                    iI = int.Parse(item.Field<decimal>("ID").ToString());
                                    sICode = item.Field<string>("Customer Code");
                                    break;
                                }
                            }

                            //iPCode
                            iP = 0;
                            var qProduct = from dt in dtProduct.AsEnumerable()//
                                           where (dt.Field<string>("Product Code") == ClassIm.GetPCode(((Range)(m_Worksheet.Cells[i, iTableSAPorder[0]])).Text.ToString()))
                                           select dt;
                            if (qProduct.Count() <= 0)
                            {
                                textBoxERRORText += m_Worksheet.Name + " Line " + i.ToString() + ": NO Product Code '" + ClassIm.GetPCode(((Range)(m_Worksheet.Cells[i, iTableSAPorder[0]])).Text.ToString()) + "'\r\n";
                                backgroundWorkerImport.ReportProgress(0);
                                continue;
                            }
                            else
                            {
                                foreach (var item in qProduct)//显示查询结果
                                {
                                    iP = int.Parse(item.Field<decimal>("ID").ToString());
                                    sPCode = item.Field<string>("Product Code");
                                    break;
                                }
                            }


                            cImeOrder ci = new cImeOrder(iP, sPCode, iI, sICode, int.Parse(((Range)(m_Worksheet.Cells[i, iTableSAPorder[3]])).Value2.ToString()), int.Parse(sT[0]), int.Parse(sT[1]));
                            alSAP.Add(ci);
                        }
                        else //存在记录，直接加入
                        {
                            foreach (cImeOrder s in alQuery)
                            {
                                s.iNum += int.Parse(((Range)(m_Worksheet.Cells[i, iTableSAPorder[3]])).Value2.ToString());
                            }
                        }

                    }
                }

                #endregion

                #region Salesplan
                if (iSalesplan > 0)
                {
                    m_Worksheet = (Microsoft.Office.Interop.Excel.Worksheet)m_Workbook.Worksheets.get_Item(iSalesplan);

                    iTemp = int.Parse(ClassIm.GetWeek(dateTimePickerP.Value, iSAPControlNumber - 1));
                    iTemp = int.Parse(ClassIm.sYear) * 100 + iTemp;

                    toolStripProgressBarALLValue++;
                    textBoxLOGText += "\r\nRead TABLE '" + m_Worksheet.Name + "'\r\n";

                    toolStripProgressBarExcelMaximum = m_Worksheet.UsedRange.Rows.Count + 1;
                    toolStripProgressBarExcelValue = 0;
                    backgroundWorkerImport.ReportProgress(0);

                    for (i = 2; i <= m_Worksheet.UsedRange.Rows.Count; i++)
                    {
                        toolStripStatusLabelWarnText = "LINE:" + i.ToString();
                        toolStripProgressBarExcelValue = i;
                        backgroundWorkerImport.ReportProgress(0);


                        //隐藏行
                        if ((bool)m_Worksheet.get_Range(m_Worksheet.Rows[i, objOpt], m_Worksheet.Rows[i, objOpt]).Hidden && !checkBoxHidden.Checked)
                            continue;

                        //iICODE Sales plan 采用客户名称连接
                        iI = 0; sICode = "";
                        var qCustomer = from dt in dtCustumor.AsEnumerable()//
                                        where (dt.Field<string>("Customer Name") == ((Range)(m_Worksheet.Cells[i, iTableSalesplan[1]])).Text.ToString().Trim())
                                        select dt;
                        if (qCustomer.Count() <= 0)
                        {
                            textBoxERRORText += m_Worksheet.Name + " Line " + i.ToString() + ": NO Customers Name '" + ((Range)(m_Worksheet.Cells[i, iTableSalesplan[1]])).Text.ToString() + "'\r\n";
                            backgroundWorkerImport.ReportProgress(0);
                            continue;
                        }
                        else
                        {
                            foreach (var item in qCustomer)//显示查询结果
                            {
                                iI = int.Parse(item.Field<decimal>("ID").ToString());
                                sICode = item.Field<string>("Customer Code");
                                break;
                            }
                        }

                        //iPCode
                        iP = 0; sPCode = "";
                        var qProduct = from dt in dtProduct.AsEnumerable()//
                                       where (dt.Field<string>("Product Code") == ClassIm.GetPCode(((Range)(m_Worksheet.Cells[i, iTableSalesplan[0]])).Text.ToString()))
                                       select dt;
                        if (qProduct.Count() <= 0)
                        {
                            textBoxERRORText += m_Worksheet.Name + " Line " + i.ToString() + ": NO Product Code '" + ClassIm.GetPCode(((Range)(m_Worksheet.Cells[i, iTableSalesplan[0]])).Text.ToString()) + "'\r\n";
                            backgroundWorkerImport.ReportProgress(0);
                            continue;
                        }
                        else
                        {
                            foreach (var item in qProduct)//显示查询结果
                            {
                                iP = int.Parse(item.Field<decimal>("ID").ToString());
                                sPCode = item.Field<string>("Product Code");
                                break;
                            }
                        }

                        //读取
                        for (j = 3; j <= m_Worksheet.UsedRange.Columns.Count; j++)
                        {
                            //跳过空行

                            if (((Range)(m_Worksheet.Cells[i, j])).Value2 == null)
                            {
                                continue;
                            }

                            //检查表头周数
                            if (((Range)(m_Worksheet.Cells[1, j])).Text.ToString() == "")
                            {
                                continue;
                            }
                            if (((Range)(m_Worksheet.Cells[1, j])).Text.ToString().Length < 7)
                            {
                                continue;
                            }

                            iYear = int.Parse(((Range)(m_Worksheet.Cells[1, j])).Text.ToString().Substring(0, 4));
                            if (((Range)(m_Worksheet.Cells[1, j])).Text.ToString().Length==8)
                                iWeek = int.Parse(((Range)(m_Worksheet.Cells[1, j])).Text.ToString().Substring(6, 2));
                            else
                                iWeek = int.Parse(((Range)(m_Worksheet.Cells[1, j])).Text.ToString().Substring(6, 1));

                            if (iYear * 100 + iWeek <= iTemp)
                            {
                                if (checkBoxDetail.Checked)
                                {
                                    textBoxLOGText += "Sale plan" + iYear.ToString() + "-" + iWeek.ToString() + " ignore \r\n";
                                    backgroundWorkerImport.ReportProgress(0);
                                }
                                //continue;
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


                //处理ORDER
                //SAP有效周记数
                toolStripProgressBarALLValue = toolStripProgressBarALL.Maximum - 2;
                textBoxLOGText += "\r\nCalculate SAP order & Sales plan:\r\n";

                iTemp = int.Parse(ClassIm.GetWeek(dateTimePickerP.Value, iSAPControlNumber - 1));
                iTemp = int.Parse(ClassIm.sYear) * 100 + iTemp;

                iTemp1 = dateTimePickerP.Value.Year * 100 + int.Parse(sNowWeek);

                toolStripProgressBarExcelMaximum = alSAP.Count + 1;
                toolStripProgressBarExcelValue = 0;
                backgroundWorkerImport.ReportProgress(0);

                foreach (cImeOrder cime in alSAP)
                {
                    toolStripProgressBarExcelValue++;
                    toolStripStatusLabelWarnText = "ITEM:" + toolStripProgressBarExcel.Value.ToString();
                    backgroundWorkerImport.ReportProgress(0);


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
                                    textBoxLOGText += ci.iYear.ToString() + "-" + ci.iWeek.ToString() + " " + ci.sICode + "|" + ci.sPCode + ":" + ci.iNum.ToString() + "=MAX(" + ci.iNum + "," + cime.iNum + ")\r\n";
                                    backgroundWorkerImport.ReportProgress(0);
                                }

                            }
                            else //在sap控制范围内
                            {
                                ci.iNum = cime.iNum; //取SAP值
                                if (checkBoxDetail.Checked)
                                {
                                    textBoxLOGText += ci.iYear.ToString() + "-" + ci.iWeek.ToString() + " " + ci.sICode + "|" + ci.sPCode + ":SAP(" + ci.iNum + ")\r\n";
                                    backgroundWorkerImport.ReportProgress(0);
                                }
                            }
                        }
                    }

                }

                //导入数据库
                toolStripProgressBarALLValue = toolStripProgressBarALL.Maximum - 1;
                textBoxLOGText += "\r\nInsert Data to DataTable:\r\n";
                toolStripProgressBarExcelMaximum = alSales.Count + 1;
                toolStripProgressBarExcelValue = 0;
                backgroundWorkerImport.ReportProgress(0);
                try
                {
                    i = 1;
                    foreach (cImeOrder cime1 in alSales)
                    {
                        toolStripProgressBarExcelValue++;
                        backgroundWorkerImport.ReportProgress(0);
                        //ScrollToCaret();
                        //this.Refresh();

                        toolStripStatusLabelWarnText = "ITEM:" + toolStripProgressBarExcel.Value.ToString();

                        var qOrder = from dt2 in dtOrder.AsEnumerable()//
                                     where (dt2.Field<string>("Product Code") == cime1.sPCode) && (dt2.Field<string>("Customer Code") == cime1.sICode) && (dt2.Field<decimal>("Year") == cime1.iYear) && (dt2.Field<string>("Num of Week") == cime1.iWeek.ToString())//条件
                                     select dt2;
                        if (qOrder.Count() <= 0)
                        {
                            oTemp6[0] = i;
                            oTemp6[1] = cime1.iPCode.ToString();
                            oTemp6[2] = cime1.sPCode;
                            oTemp6[3] = cime1.iICode.ToString();
                            oTemp6[4] = cime1.sICode;
                            oTemp6[5] = cime1.iYear;
                            oTemp6[6] = cime1.iWeek.ToString();
                            oTemp6[7] = cime1.iNum.ToString();
                            oTemp6[8] = 0;
                            i++;
                            dtOrder.Rows.Add(oTemp6);


                        }


                    }


                }
                catch (Exception ex)
                {
                    MessageBox.Show("error：" + ex.Message.ToString(), "error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                finally
                {

                }

                #endregion

            } //try input
            catch (Exception exc)
            {
                MessageBox.Show("Input Fail：" + exc.Message.ToString(), "Error");
            }
            finally
            {
                //

                dataGridViewProduct.DataSource = dtProduct;
                dataGridViewCustomer.DataSource = dtCustumor;
                dataGridViewTAC.DataSource = dtTAC;
                dataGridViewReserve.DataSource = dtAcquire;
                dataGridViewCW.DataSource = dtActual;
                dataGridViewBacklog.DataSource = dtBacklog;
                dataGridViewOrder.DataSource = dtOrder;

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

        private void FormImportDataOffLine_FormClosed(object sender, FormClosedEventArgs e)
        {
            backgroundWorkerImport.CancelAsync();
        }

        private void backgroundWorkerImport_RunWorkerCompleted(object sender, System.ComponentModel.RunWorkerCompletedEventArgs e)
        {
            if (e.Error != null)
            {
                //MessageBox.Show("Error");
            }
            else if (e.Cancelled)
            {
                //MessageBox.Show("Canceled");
            }
            else
            {
                btnLoad.Enabled = true;
                toolStripProgressBarExcel.Value = toolStripProgressBarExcel.Maximum;
                toolStripProgressBarALL.Value = toolStripProgressBarALL.Maximum;
                toolStripStatusLabelWarn.Text = "";
                if (MessageBox.Show("Input Finish,do you want to open forecaset window?", "Information", MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.Yes)
                {
                    FormManageOffline childFormManageOffline = new FormManageOffline();
                    // 在显示该窗体前使其成为此 MDI 窗体的子窗体。
                    childFormManageOffline.MdiParent = f;

                    childFormManageOffline.Text = "Forecast Offline : "+ExcelPathName;

                    childFormManageOffline.dtProduct = dtProduct;
                    childFormManageOffline.dtCustumor = dtCustumor;
                    childFormManageOffline.dtTAC = dtTAC;
                    childFormManageOffline.dtAcquire = dtAcquire;
                    childFormManageOffline.dtActual = dtActual;
                    childFormManageOffline.dtOrder = dtOrder;
                    childFormManageOffline.dtBacklog = dtBacklog;

                    childFormManageOffline.iNumofPass = iNumofPass;
                    childFormManageOffline.iNumofWeek = iNumofWeek;

                    childFormManageOffline.Show();
                    //this.Close();
                }
            }
        }

        private void backgroundWorkerImport_ProgressChanged(object sender, System.ComponentModel.ProgressChangedEventArgs e)
        {
            textBoxERROR.Text = textBoxERRORText;
            textBoxLOG.Text = textBoxLOGText;

            toolStripProgressBarALL.Maximum = toolStripProgressBarALLMaximum;
            toolStripProgressBarALL.Value = toolStripProgressBarALLValue;

            toolStripProgressBarExcel.Maximum = toolStripProgressBarExcelMaximum;
            toolStripProgressBarExcel.Value = toolStripProgressBarExcelValue;

            toolStripStatusLabelWarn.Text = toolStripStatusLabelWarnText;

            ScrollToCaret();

        }


    }

    
}
