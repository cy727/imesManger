using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Collections;

namespace imesManger
{


    public partial class FormManage : Form
    {
        private System.Data.SqlClient.SqlConnection sqlConn = new System.Data.SqlClient.SqlConnection();
        private System.Data.SqlClient.SqlCommand sqlComm = new System.Data.SqlClient.SqlCommand();
        private System.Data.SqlClient.SqlDataReader sqldr;
        private System.Data.SqlClient.SqlDataAdapter sqlDA = new System.Data.SqlClient.SqlDataAdapter();
        private System.Data.DataSet dSet = new DataSet();

        public string strConn = "";

        public int intUserLimit = 0;
        private DataTable dtOrder = new DataTable();

        public int iNumofWeek = 0;
        public int iNumofPass = 0;

        private ArrayList alIme = new ArrayList();
        public int iRange = 1000000;

        private const int ICOLUMNS = 16;
        private const char SPLIT = ':';

        public FormManage()
        {
            InitializeComponent();
        }

        private void FormManage_Load(object sender, EventArgs e)
        {
            sqlConn.ConnectionString = strConn;
            sqlComm.Connection = sqlConn;
            sqlDA.SelectCommand = sqlComm;

            this.Top = 1;
            this.Left = 1;

            initDatatable(false,false);

        }

        private void setSTAUS()
        {
            toolStripStatusLabelC.Text = "Number of ORDER Rows: " + dataGridViewP.RowCount.ToString();
        }

        private void initDatatable(bool bPCode, bool bICode)
        {
            int i,j,k;
            int iTemp = 0, iTemp1=0;
            string sTemp = "";
            bool bFirst = true;

            dataGridViewP.RowValidating -= dataGridViewP_RowValidating;
            dataGridViewP.CellValidating -= dataGridViewP_CellValidating;

            dtOrder.Columns.Clear();
            dtOrder.Columns.Add("pid", System.Type.GetType("System.Decimal"));//0
            dtOrder.Columns.Add("Product Code", System.Type.GetType("System.String"));
            dtOrder.Columns.Add("Product Name", System.Type.GetType("System.String"));
            dtOrder.Columns.Add("SIM Slot", System.Type.GetType("System.Decimal"));
            dtOrder.Columns.Add("Failure Rate", System.Type.GetType("System.Decimal"));
            dtOrder.Columns.Add("iid", System.Type.GetType("System.Decimal"));//4
            dtOrder.Columns.Add("Customer Code", System.Type.GetType("System.String"));
            dtOrder.Columns.Add("TAC ID", System.Type.GetType("System.Decimal"));
            dtOrder.Columns.Add("TAC Code", System.Type.GetType("System.String"));
            dtOrder.Columns.Add("TAC Init", System.Type.GetType("System.Decimal"));
            dtOrder.Columns.Add("Manufacture Number", System.Type.GetType("System.Decimal"));//9
            dtOrder.Columns.Add("Used IMEI Quantity", System.Type.GetType("System.Decimal"));
            dtOrder.Columns.Add(" Production week", System.Type.GetType("System.String"));
            dtOrder.Columns.Add("Acquire Number", System.Type.GetType("System.Decimal"));
            dtOrder.Columns.Add("Acquire Date", System.Type.GetType("System.String"));
            dtOrder.Columns.Add("Remaining", System.Type.GetType("System.Decimal")); //剩余 14

            for (i = iNumofPass; i > 0; i--)
            {
                sTemp = ClassIm.GetWeek(dateTimePickerP.Value, -1 * i);
                sTemp = ClassIm.sYear + "-" + sTemp;

                dtOrder.Columns.Add(sTemp, System.Type.GetType("System.String"));
            }

            for (i = 0; i < iNumofWeek; i++)
            {
                sTemp = ClassIm.GetWeek(dateTimePickerP.Value, i);
                sTemp = ClassIm.sYear + "-" + sTemp;

                dtOrder.Columns.Add(sTemp, System.Type.GetType("System.String"));
            }


            sqlConn.Open();
            bFirst=true;
            sqlComm.CommandText = "SELECT product.ID, product.[Product Name], product.[Product Code], indentor.ID AS iID, indentor.[Indentor Name],indentor.[Indentor Code], product.[Number of IMEI], product.[Failure Rate] FROM product CROSS JOIN indentor ";
            if (bPCode && textBoxPC.Text.Trim()!="")
            {
                sqlComm.CommandText += " WHERE (product.[Product Code] LIKE N'%" + textBoxPC.Text.Trim() + "%')";
                bFirst = false;
            }
            if (bICode && textBoxIC.Text.Trim() != "")
            {
                if(bFirst)
                    sqlComm.CommandText += " WHERE (indentor.[Indentor Code] LIKE N'%" + textBoxIC.Text.Trim() + "%')";
                else
                    sqlComm.CommandText += " AND (indentor.[Indentor Code] LIKE N'%" + textBoxIC.Text.Trim() + "%')";
            }
            if (dSet.Tables.Contains("pi")) dSet.Tables["pi"].Clear();
            sqlDA.Fill(dSet, "pi");

            bFirst = true;
            sqlComm.CommandText = "SELECT ID, [TAC Code], [Product ID], [Product Code], [Indentor ID], [Indentor Code], [Init Number] FROM      TAC";
            if (bPCode && textBoxPC.Text.Trim() != "")
            {
                sqlComm.CommandText += " WHERE (TAC.[Product Code] LIKE N'%" + textBoxPC.Text.Trim() + "%')";
                bFirst = false;
            }
            if (bICode && textBoxIC.Text.Trim() != "")
            {
                if (bFirst)
                    sqlComm.CommandText += " WHERE (TAC.[Indentor Code] LIKE N'%" + textBoxIC.Text.Trim() + "%')";
                else
                    sqlComm.CommandText += " AND (TAC.[Indentor Code] LIKE N'%" + textBoxIC.Text.Trim() + "%')";
            }
            sqlComm.CommandText += " ORDER BY [Init Number]";
            if (dSet.Tables.Contains("TAC")) dSet.Tables["TAC"].Clear();
            sqlDA.Fill(dSet, "TAC");



            bFirst = true;
            sqlComm.CommandText = "SELECT  ID, [Product ID], [Product Code], [Indentor ID], [Indentor Code], [End Number], Year, [Num of Week], [Start Number], Date, [TAC ID], [TAC Code], Status FROM acquire";
            if (bPCode && textBoxPC.Text.Trim() != "")
            {
                sqlComm.CommandText += " WHERE ([Product Code] LIKE N'%" + textBoxPC.Text.Trim() + "%')";
                bFirst = false;
            }
            if (bICode && textBoxIC.Text.Trim() != "")
            {
                if (bFirst)
                    sqlComm.CommandText += " WHERE ([Indentor Code] LIKE N'%" + textBoxIC.Text.Trim() + "%')";
                else
                    sqlComm.CommandText += " AND ([Indentor Code] LIKE N'%" + textBoxIC.Text.Trim() + "%')";
            }
            if (dSet.Tables.Contains("acquire")) dSet.Tables["acquire"].Clear();
            sqlDA.Fill(dSet, "acquire");

            bFirst = true;
            sqlComm.CommandText = "SELECT  actual_1.[Product ID], actual_1.[Product Code], actual_1.[Indentor ID], actual_1.[Indentor Code], actual_1.[Start Number], actual_1.[End Number], actual_1.[Acquire ID], actual_1.Date, actual_1.Year, actual_1.[Num of Week], actual_1.[TAC ID], actual_1.[TAC Code], actual_1.[Total Number] FROM actual INNER JOIN actual AS actual_1 ON actual.[Product Code] = actual_1.[Product Code] AND actual.[Indentor Code] = actual_1.[Indentor Code] GROUP BY actual_1.[Product ID], actual_1.[Product Code], actual_1.[Indentor ID], actual_1.[Indentor Code], actual_1.[Start Number], actual_1.[End Number], actual_1.[Acquire ID], actual_1.Date, actual_1.Year, actual_1.[Num of Week], actual_1.[TAC ID], actual_1.[TAC Code], actual_1.[Total Number]";
            if (bPCode && textBoxPC.Text.Trim() != "")
            {
                sqlComm.CommandText += " HAVING (actual_1.[Product Code] LIKE N'%" + textBoxPC.Text.Trim() + "%')";
                bFirst = false;
            }
            if (bICode && textBoxIC.Text.Trim() != "")
            {
                if (bFirst)
                    sqlComm.CommandText += " HAVING (actual_1.[Indentor Code] LIKE N'%" + textBoxIC.Text.Trim() + "%')";
                else
                    sqlComm.CommandText += " AND (actual_1.[Indentor Code] LIKE N'%" + textBoxIC.Text.Trim() + "%')";
            }
            if (dSet.Tables.Contains("actual")) dSet.Tables["actual"].Clear();
            sqlDA.Fill(dSet, "actual");

            bFirst = true;
            sqlComm.CommandText = "SELECT   ID, [Product ID], [Product Code], [Indentor ID], [Indentor Code], Year, [Num of Week], Numbers, backlog FROM orders ";
            if (bPCode && textBoxPC.Text.Trim() != "")
            {
                sqlComm.CommandText += " WHERE ([Product Code] LIKE N'%" + textBoxPC.Text.Trim() + "%')";
                bFirst = false;
            }
            if (bICode && textBoxIC.Text.Trim() != "")
            {
                if (bFirst)
                    sqlComm.CommandText += " WHERE ([Indentor Code] LIKE N'%" + textBoxIC.Text.Trim() + "%')";
                else
                    sqlComm.CommandText += " AND ([Indentor Code] LIKE N'%" + textBoxIC.Text.Trim() + "%')";
            }
            if (dSet.Tables.Contains("orders")) dSet.Tables["orders"].Clear();
            sqlDA.Fill(dSet, "orders");


            sqlConn.Close();

            object[] oTemp = new object[ICOLUMNS+iNumofWeek+iNumofPass];
            dtOrder.Clear();

 
            for (i = 0; i < dSet.Tables["pi"].Rows.Count; i++)
            {
                oTemp[0] = dSet.Tables["pi"].Rows[i][0];
                oTemp[1] = dSet.Tables["pi"].Rows[i][2];
                oTemp[2] = dSet.Tables["pi"].Rows[i][1];
                oTemp[3] = dSet.Tables["pi"].Rows[i][6];
                oTemp[4] = dSet.Tables["pi"].Rows[i][7];
                oTemp[5] = dSet.Tables["pi"].Rows[i][3];
                oTemp[6] = dSet.Tables["pi"].Rows[i][5];

                //TAC Manufacture
                oTemp[7] = 0;
                oTemp[8] = "";
                oTemp[9] = 0;
                oTemp[10] = 0;
                oTemp[11] = 0;
                oTemp[12] = "2000-1";

                var q2 = from dt1 in dSet.Tables["actual"].AsEnumerable()//查询最后
                         where (dt1.Field<string>("Product Code") == oTemp[1].ToString()) && (dt1.Field<string>("Indentor Code") == oTemp[6].ToString())//条件
                         select dt1;
                
                foreach (var item in q2)//显示查询结果
                {
                    //TAC
                    oTemp[7] = item.Field<int>("TAC ID");
                    oTemp[8] = item.Field<string>("TAC Code").ToString();

                    var qTAC = from dt2 in dSet.Tables["TAC"].AsEnumerable()//得到TAC初始值
                             where (dt2.Field<string>("Product Code") == oTemp[1].ToString()) && (dt2.Field<string>("Indentor Code") == oTemp[6].ToString()) && (dt2.Field<string>("TAC Code") == oTemp[8].ToString())//条件
                             select dt2;
                    foreach (var item1 in qTAC)//显示查询结果
                    {
                        oTemp[9] = item1.Field<int>("Init Number");
                        break;
                    }

                    //Manufacture
                    oTemp[10] = item.Field<int>("End Number");
                    oTemp[11] = item.Field<int>("Total Number");
                    oTemp[12] = item.Field<int>("Year").ToString() + "-" + item.Field<string>("Num of Week");
                    break;
                }

                //acquire
                oTemp[13] = 0;
                oTemp[14] = "";
                var q3 = from dt1 in dSet.Tables["acquire"].AsEnumerable()//查询最后
                         where (dt1.Field<string>("Product Code") == oTemp[1].ToString()) && (dt1.Field<string>("Indentor Code") == oTemp[6].ToString())//条件
                         orderby dt1.Field<int>("End Number") descending
                         select dt1;

                foreach (var item in q3)//显示查询结果
                {
                    oTemp[13] = item.Field<int>("End Number");
                    oTemp[14] = item.Field<int>("Year").ToString() + "-" + item.Field<string>("Num of Week");
                    break;
                }

                //剩余号码
                oTemp[15] = getSurplus(oTemp[1].ToString(), oTemp[6].ToString(), int.Parse(oTemp[10].ToString()));


                //ORDER numbers
                k = ICOLUMNS;
                for (j = iNumofPass; j > 0; j--)
                {
                    oTemp[k] = 0;
                    //得到日期
                    string[] sT = dtOrder.Columns[k].Caption.Split('-');
                    if (sT.Length == 2)
                    {
                        var q4 = from dt1 in dSet.Tables["orders"].AsEnumerable()//查询最后
                                 where (dt1.Field<string>("Product Code") == oTemp[1].ToString()) && (dt1.Field<string>("Indentor Code") == oTemp[6].ToString()) && (dt1.Field<int>("Year") == int.Parse(sT[0])) && (dt1.Field<string>("Num of Week") == sT[1])//条件
                                 select dt1;
                        foreach (var item in q4)//显示查询结果
                        {
                            oTemp[k] = item.Field<int>("Numbers");
                            break;
                        }
                    }
                    k++;
                }

                for (j = 0; j < iNumofWeek; j++)
                {
                    oTemp[k] = 0;
                    //得到日期
                    string[] sT = dtOrder.Columns[k].Caption.Split('-');
                    if (sT.Length == 2)
                    {
                        var q5 = from dt1 in dSet.Tables["orders"].AsEnumerable()//查询最后
                                 where (dt1.Field<string>("Product Code") == oTemp[1].ToString()) && (dt1.Field<string>("Indentor Code") == oTemp[6].ToString()) && (dt1.Field<int>("Year") == int.Parse(sT[0])) && (dt1.Field<string>("Num of Week") == sT[1])//条件
                                 select dt1;
                        foreach (var item in q5)//显示查询结果
                        {
                            oTemp[k] = item.Field<int>("Numbers");
                            break;
                        }
                    }
                    k++;
                }


               dtOrder.Rows.Add(oTemp);
            }

            dataGridViewP.DataSource = dtOrder;
            for (i = 0; i < dataGridViewP.Columns.Count; i++)
            {
                dataGridViewP.Columns[i].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                dataGridViewP.Columns[i].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            }
            for (i = 0; i < dataGridViewP.Columns.Count; i++)
            {
                dataGridViewP.Columns[i].ReadOnly = true;
            }
            dataGridViewP.Columns[0].Visible = false;
            dataGridViewP.Columns[5].Visible = false;
            dataGridViewP.Columns[7].Visible = false;
            dataGridViewP.Columns[9].Visible = false;
            dataGridViewP.Columns[10].Visible = false;
            //dataGridViewP.Columns[0].Visible = false;
            //dataGridViewP.Columns[0].Visible = false;
            
            setSTAUS();

            //低于实际时间不可编辑
            //for (i = 0; i < dataGridViewP.RowCount; i++)
            //{
            //    string[] sT = dataGridViewP.Rows[i].Cells[7].Value.ToString().Split('-');
            //    if(sT.Length==2)
            //    {
            //        iTemp = int.Parse(sT[0]) * 100 + int.Parse(sT[1]);

            //        for (j = 10; j < dataGridViewP.ColumnCount; j++) //判断
            //        {
            //            string[] sT1 = dataGridViewP.Columns[j].Name.Split('-');
            //            iTemp1 = int.Parse(sT1[0]) * 100 + int.Parse(sT1[1]);

            //            if (iTemp1 >= iTemp)
            //                dataGridViewP.Rows[i].Cells[j].ReadOnly = false;
            //            else
            //                dataGridViewP.Rows[i].Cells[j].ReadOnly = true;

            //        }
            //    }
            //}



            //make arraylist
            alIme.Clear();

            for (i = 0; i < dataGridViewP.RowCount; i++)
            {
                string[] sT = dataGridViewP.Rows[i].Cells[12].Value.ToString().Split('-');
                iTemp=int.Parse(sT[0]) * 100 + int.Parse(sT[1]);
                var q10 = from dt1 in dSet.Tables["orders"].AsEnumerable()//大于起始时间Order
                          where (dt1.Field<string>("Product Code") == dataGridViewP.Rows[i].Cells[1].Value.ToString()) && (dt1.Field<string>("Indentor Code") == dataGridViewP.Rows[i].Cells[6].Value.ToString()) && (dt1.Field<int>("Year") * 100 + int.Parse(dt1.Field<string>("Num of Week"))>=iTemp)
                         orderby (dt1.Field<int>("Year") * 100 + int.Parse(dt1.Field<string>("Num of Week")))  
                         select dt1;
                foreach (var item in q10)//显示查询结果
                {
                    cIme cime = new cIme(int.Parse(dataGridViewP.Rows[i].Cells[0].Value.ToString()),dataGridViewP.Rows[i].Cells[1].Value.ToString(), int.Parse(dataGridViewP.Rows[i].Cells[5].Value.ToString()),dataGridViewP.Rows[i].Cells[6].Value.ToString(), int.Parse(dataGridViewP.Rows[i].Cells[3].Value.ToString()));
                    cime.iYear = item.Field<int>("Year");
                    cime.iWeek = int.Parse(item.Field<string>("Num of Week"));

                    cime.iCount = item.Field<int>("Numbers");
                    cime.iBacklog = item.Field<int>("backlog");

                    cime.dFailureRate = decimal.Parse(dataGridViewP.Rows[i].Cells[4].Value.ToString());

                    alIme.Add(cime);
                }

                //加入未有数据
                for (j = ICOLUMNS; j < dataGridViewP.Columns.Count; j++)
                {
                    //判断是有效日期
                    string[] sT1 = dataGridViewP.Columns[j].Name.Split('-');
                    iTemp1=int.Parse(sT1[0]) * 100 + int.Parse(sT1[1]);
                    if (iTemp1 <= iTemp) continue;

                    var alQuery = from cIme cime in alIme
                                  where (cime.sPCode == dataGridViewP.Rows[i].Cells[1].Value.ToString()) && (cime.sICode == dataGridViewP.Rows[i].Cells[6].Value.ToString()) && (cime.iYear == int.Parse(sT1[0])) && (cime.iWeek == int.Parse(sT1[1]))
                                  select cime;

                    //foreach (cIme s in alQuery)
                    //{
                    //}
                    if (alQuery.Count() < 1) //数据库没有记录，加入预订为0的数据
                    {
                        cIme cime1 = new cIme(int.Parse(dataGridViewP.Rows[i].Cells[0].Value.ToString()), dataGridViewP.Rows[i].Cells[1].Value.ToString(), int.Parse(dataGridViewP.Rows[i].Cells[5].Value.ToString()), dataGridViewP.Rows[i].Cells[6].Value.ToString(), int.Parse(dataGridViewP.Rows[i].Cells[3].Value.ToString()));
                        cime1.iYear = int.Parse(sT1[0]);
                        cime1.iWeek = int.Parse(sT1[1]);

                        cime1.iCount = 0;
                        cime1.dFailureRate = decimal.Parse(dataGridViewP.Rows[i].Cells[4].Value.ToString());
                        alIme.Add(cime1);
                    }
                }


                
            }
            //manageArrayList();
            countAmount(-1);
            dataGridViewP.RowValidating += dataGridViewP_RowValidating;
            dataGridViewP.CellValidating += dataGridViewP_CellValidating;
        }

        //取得当前剩余
        private int getSurplus(string sPCode,string sICode,int iManNumber)
        {
            if (sPCode.Trim() == "" || sICode=="")
                return 0;

            int i;
            int iSurplus = 0;
            var qAcquire = from dt1 in dSet.Tables["acquire"].AsEnumerable()//查询最后
                     where (dt1.Field<string>("Product Code") == sPCode) && (dt1.Field<string>("Indentor Code") == sICode)//条件
                     orderby dt1.Field<int>("End Number") ascending
                     select dt1;

            foreach (var item in qAcquire)//显示查询结果
            {
                if (iManNumber > item.Field<int>("End Number"))
                    continue;

                if (iManNumber < item.Field<int>("Start Number"))
                {
                    iSurplus += item.Field<int>("End Number") - item.Field<int>("Start Number");
                    continue;
                }
                if (iManNumber >= item.Field<int>("Start Number") && iManNumber <= item.Field<int>("End Number"))
                {
                    iSurplus += item.Field<int>("End Number") - iManNumber;
                    continue;
                }

            }

            return iSurplus;
        }

        private bool manageArrayList()
        {
            bool bResult = true;
            int i, j, k;


            //调整al=用户输入数量
            for (i = 0; i < dataGridViewP.RowCount; i++)
            {
                for (j = ICOLUMNS; j < dataGridViewP.ColumnCount; j++)
                {
                    string[] sT1 = dataGridViewP.Columns[j].Name.Split('-');
                    var alQuery = from cIme cime in alIme
                                  where (cime.sPCode == dataGridViewP.Rows[i].Cells[1].Value.ToString()) && (cime.sICode == dataGridViewP.Rows[i].Cells[6].Value.ToString()) && (cime.iYear == int.Parse(sT1[0])) && (cime.iWeek == int.Parse(sT1[1]))
                                  select cime;
                    foreach (cIme s in alQuery)
                    {
                        if (dataGridViewP.Rows[i].Cells[j].Value.ToString().IndexOf(SPLIT) == -1)
                            s.iCount = int.Parse(dataGridViewP.Rows[i].Cells[j].Value.ToString());
                        else
                        {
                            string[] sST = dataGridViewP.Rows[i].Cells[j].Value.ToString().Split(SPLIT);
                            s.iCount = int.Parse(sST[0]);
                        }
                    }
                }

            }
            return bResult;
        }

        private void btnRe_Click(object sender, EventArgs e)
        {
            initDatatable(false,false);
        }

        private void toolStripButtonEDIT_Click(object sender, EventArgs e)
        {
            int i;
            int iTemp;

            if (!countAmount(-1))
            {
                MessageBox.Show("please check data", "information", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            System.Data.SqlClient.SqlTransaction sqlta;
            sqlConn.Open();
            sqlta = sqlConn.BeginTransaction();
            sqlComm.Transaction = sqlta;
            try
            {
                foreach (cIme s in alIme)
                {
                    sqlComm.CommandText = "SELECT ID, [Product ID], [Product Code], [Indentor ID], [Indentor Code], [Start Number], [End Number], [Acquire ID], Date,  Year, [Num of Week], Numbers FROM orders WHERE ([Product Code] = N'"+s.sPCode+"') AND ([Indentor Code] = N'"+s.sICode+"') AND (Year = "+s.iYear.ToString()+") AND ([Num of Week] = N'"+s.iWeek.ToString()+"')";

                    sqldr = sqlComm.ExecuteReader();
                    if (sqldr.HasRows)
                    {
                        sqldr.Read();
                        iTemp=int.Parse(sqldr.GetValue(0).ToString());
                        sqldr.Close();
                        sqlComm.CommandText = "UPDATE  orders SET [Start Number] = " + s.iStart.ToString() + ", [End Number] = " + s.iEnd.ToString() + ", [Acquire ID] = " + s.iAcquire.ToString() + ", Numbers = " + s.iCount.ToString() + ", [Order Start Number] =" + s.iOrderStart.ToString() + ", [Order End Number] =" + s.iOrderEnd.ToString() + " WHERE (ID = " + iTemp.ToString() + ")";
                    }
                    else
                    {
                        sqldr.Close();
                        DateTime dt = ClassIm.GetDate(s.iYear, s.iWeek);
                        sqlComm.CommandText = "INSERT INTO orders ([Product ID], [Product Code], [Indentor ID], [Indentor Code], [Start Number], [End Number], [Acquire ID], Date, Year,  [Num of Week], Numbers,[Order Start Number],[Order End Number],backlog) VALUES  (" + s.iPCode.ToString() + ", N'" + s.sPCode + "', " + s.iICode.ToString() + ", N'" + s.sICode + "', " + s.iStart.ToString() + ", " + s.iEnd.ToString() + ", " + s.iAcquire.ToString() + ", CONVERT(DATETIME, '" + dt.ToShortDateString() + " 00:00:00', 102), " + s.iYear.ToString() + ", N'" + s.iWeek.ToString() + "', " + s.iCount.ToString() + "," + s.iOrderStart.ToString() + "," + s.iOrderEnd.ToString() + ",0)";
                    }
                    sqlComm.ExecuteNonQuery();
                }
                sqlta.Commit();
                MessageBox.Show("ORDER finished");
            }
            catch (Exception ex)
            {
                MessageBox.Show("error：" + ex.Message.ToString(), "error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                sqlta.Rollback();
                return;
            }
            finally
            {
                sqlConn.Close();
            }
        }

        private void dataGridViewP_CellValidating(object sender, DataGridViewCellValidatingEventArgs e)
        {
            //if (dataGridViewP.Rows[e.RowIndex].IsNewRow)
            //    return;

            //ClassIm.ClearDataGridViewErrorText(dataGridViewP);
            //if(e.ColumnIndex>10)
            //{
            //        int intOut = 0;
            //        if (e.FormattedValue.ToString() == "") return;
            //        if (int.TryParse(e.FormattedValue.ToString(), out intOut))
            //        {
            //            if (intOut < 0)
            //            {
            //                dataGridViewP.Rows[e.RowIndex].Cells[e.ColumnIndex].ErrorText = "data error";
            //                e.Cancel = true;
            //            }
            //        }
            //        else
            //        {
            //            dataGridViewP.Rows[e.RowIndex].Cells[e.ColumnIndex].ErrorText = "data error";
            //            e.Cancel = true;
            //        }
            //}
            //dataGridViewP.EndEdit();

        }

        private void dataGridViewP_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            MessageBox.Show("dataformat error", "information", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        private void dataGridViewP_RowValidating(object sender, DataGridViewCellCancelEventArgs e)
        {
            if (dataGridViewP.Rows.Count < 1)
                return;
            //countAmount();
        }

        private bool countAmount(int iRows)
        {
            bool bCheck = true;
            bool bFirst = true;
            int i, j, k;
            int iStartYear = 0, iStartWeek=0;
            int iEndYear = 0, iEndWeek = 0;
            int iStartNo = 0, iEndNo = 0;

            int iRStart = 0,iRend = 0;

            int iAcquireNO = 0;
            int iTemp=0;
            bool bTrue = true;

            int iSurplus = 0;
                
            //manageArrayList(); //重置用户输入
            this.dataGridViewP.CellValidating -= dataGridViewP_CellValidating;
            ClassIm.ClearDataGridViewErrorText(dataGridViewP);

            if (iRows < 0)
            {
                iRStart = 0;
                iRend = dataGridViewP.Rows.Count;
            }
            else
            {
                iRStart = iRows;
                iRend = iRows+1;
            }

            for (i = iRStart; i < iRend; i++)
            {
                if (dataGridViewP.Rows[i].IsNewRow)
                    continue;
                iSurplus = 0;
                for (j = ICOLUMNS; j < dataGridViewP.ColumnCount; j++)
                {
                    if (dataGridViewP.Rows[i].Cells[j].Value.ToString() == "")
                    {
                        dataGridViewP.Rows[i].Cells[j].ErrorText = "please input number";
                        bCheck = false;
                    }
                }

                if (!bCheck)
                    continue;

                //整理
                string[] sT = dataGridViewP.Rows[i].Cells[12].Value.ToString().Split('-');
                iStartYear = int.Parse(sT[0]); iStartWeek = int.Parse(sT[1]);
                iStartNo = int.Parse(dataGridViewP.Rows[i].Cells[11].Value.ToString());
                if(iStartNo!=0)
                       iStartNo++;//起始号


                //本周计算,backlog
                iEndWeek = iStartWeek;
                iEndYear = iStartYear;
                var alQuery1 = from cIme cime in alIme
                where (cime.sPCode == dataGridViewP.Rows[i].Cells[1].Value.ToString()) && (cime.sICode == dataGridViewP.Rows[i].Cells[6].Value.ToString()) && (cime.iYear == iEndYear) && (cime.iWeek == iEndWeek)
                select cime;
                foreach (cIme s in alQuery1)
                {
                    iAcquireNO = 0;
                    bTrue = true; 
                    while (bTrue)
                    {
                        //找到开始号区间
                        var q3 = from dt1 in dSet.Tables["acquire"].AsEnumerable()//查询最后
                                 where (dt1.Field<string>("Product Code") == s.sPCode) && (dt1.Field<string>("Indentor Code") == s.sICode) && (dt1.Field<int>("Start Number") <= iStartNo) && (dt1.Field<int>("End Number") >= iStartNo)//条件
                                 orderby dt1.Field<int>("End Number")
                                 select dt1;

                        if (q3.Count() < 1)  //没有区间
                        {
                            s.iAcquire = 0;
                            break;
                        }
                        bFirst = true; s.iAcquire = 0;
                        foreach (var item in q3)
                        {
                            if (!bFirst) //第二次循环，以此次申请的第一位数为起始数
                            {
                                iStartNo = item.Field<int>("Start Number");
                            }
                            iEndNo = iStartNo + (int)Math.Ceiling((double)s.iBacklog * (double)s.iNum * ((double)s.dFailureRate / 100.0 + 1.0)) - 1;

                            if (iEndNo <= item.Field<int>("End Number")) //找到区间
                            {
                                s.iAcquire = item.Field<int>("ID");
                                s.iAcStart = item.Field<int>("Start Number");
                                s.iAcEnd = item.Field<int>("End Number");

                                s.iStart = iStartNo; s.iEnd = iEndNo;

                                iStartNo = iEndNo + 1;
                                s.iSurpLus = getSurplus(s.sPCode, s.sICode, iEndNo);
                                bTrue = false;
                                break;
                            }
                            else //区间不够，跳到下一区间
                            {
                                //得到下一个申请区间
                                bFirst = false;
                            }
                            break;
                            
                        }
                    } //while
                    //得到预计计数
                    if (s.iAcquire == 0)
                    {
                        s.iOrderStart = iStartNo;
                        s.iOrderEnd = iStartNo + (int)Math.Ceiling((double)s.iBacklog * (double)s.iNum * ((double)s.dFailureRate / 100.0 + 1.0)) - 1;
                        iSurplus = -1 * (int)Math.Ceiling((double)s.iBacklog * (double)s.iNum * ((double)s.dFailureRate / 100.0 + 1.0));
                        s.iSurpLus = iSurplus;
                        iStartNo = s.iOrderEnd + 1;
                    }
                    break;
                }//s

                //计算以后周
                k=1;
                while (true)
                {
                    //从制造下一周开始
                    iEndWeek = int.Parse(ClassIm.GetWeek(iStartYear, iStartWeek, k));
                    iEndYear = int.Parse(ClassIm.sYear);
                    k++;
                    var alQuery = from cIme cime in alIme
                                  where (cime.sPCode == dataGridViewP.Rows[i].Cells[1].Value.ToString()) && (cime.sICode == dataGridViewP.Rows[i].Cells[6].Value.ToString()) && (cime.iYear == iEndYear) && (cime.iWeek == iEndWeek)
                                  select cime;

                    if (alQuery.Count() < 1) //没有记录
                        break;

                    foreach (cIme s in alQuery)
                    {
                        iAcquireNO = 0;

                        bTrue = true; 
                        while (bTrue)
                        {
                            //找到开始号区间
                            var q3 = from dt1 in dSet.Tables["acquire"].AsEnumerable()//查询最后
                                     where (dt1.Field<string>("Product Code") == s.sPCode) && (dt1.Field<string>("Indentor Code") == s.sICode) && (dt1.Field<int>("Start Number") <= iStartNo) && (dt1.Field<int>("End Number") >= iStartNo)//条件
                                     orderby dt1.Field<int>("End Number")
                                     select dt1;

                            if (q3.Count() < 1)  //没有区间
                            {
                                s.iAcquire = 0;                              
                                break;
                            }
                            bFirst = true; s.iAcquire = 0;
                            foreach (var item in q3)
                            {
                                if (!bFirst) //第二次循环，以此次申请的第一位数为起始数
                                {
                                    iStartNo = item.Field<int>("Start Number");
                                }
                                iEndNo = iStartNo + (int)Math.Ceiling((double)s.iCount * (double)s.iNum * ((double)s.dFailureRate / 100.0 + 1.0)) - 1;


                                if (iEndNo <= item.Field<int>("End Number")) //找到区间
                                {
                                    s.iAcquire = item.Field<int>("ID");
                                    s.iAcStart = item.Field<int>("Start Number");
                                    s.iAcEnd = item.Field<int>("End Number");

                                    s.iStart = iStartNo; s.iEnd = iEndNo;
                                    iStartNo = iEndNo + 1;
                                    s.iSurpLus = getSurplus(s.sPCode, s.sICode, iEndNo);
                                    bTrue = false;
                                    break;
                                }
                                else //区间不够，跳到下一区间
                                {
                                    bFirst=false;
                                }
                            }
                            break;
                        } //while
                        //得到预计计数
                        if (s.iAcquire == 0)
                        {
                            s.iOrderStart = iStartNo;
                            s.iOrderEnd = iStartNo + (int)Math.Ceiling((double)s.iCount * (double)s.iNum * ((double)s.dFailureRate / 100.0 + 1.0)) - 1;
                            iSurplus += -1 * ((int)Math.Ceiling((double)s.iCount * (double)s.iNum * ((double)s.dFailureRate / 100.0 + 1.0)));
                            s.iSurpLus = iSurplus;
                            iStartNo = s.iOrderEnd + 1;
                        }
                        break;
                    }//s

                }

                //if (!bCheck1)
                //    dataGridViewP.Rows[i].DefaultCellStyle.BackColor = Color.Pink;
                //else
                //    dataGridViewP.Rows[i].DefaultCellStyle.BackColor = Color.Gray;
                    //dataGridViewP.Rows[i].DefaultCellStyle.BackColor = Color.Red;

                

            } //for

            //changeViewColor();
            changeViewText();

            this.dataGridViewP.CellValidating += dataGridViewP_CellValidating;
            //dataGridViewP.EndEdit();
            

            return bCheck;
        }

        private void changeViewText()
        {
            int i, j;

            for (i = 0; i < dtOrder.Rows.Count; i++)
            {
                for (j = ICOLUMNS; j < dtOrder.Columns.Count; j++)
                {
                    string[] sT = dtOrder.Columns[j].ColumnName.Split('-');

                    var alQuery = from cIme cime in alIme
                                  where (cime.sPCode == dtOrder.Rows[i][1].ToString()) && (cime.sICode == dtOrder.Rows[i][6].ToString()) && (cime.iYear == int.Parse(sT[0])) && (cime.iWeek == int.Parse(sT[1]))
                                  select cime;
                    if (alQuery.Count() <= 0)
                        continue;
                    foreach (cIme s in alQuery)
                    {
                        //最终日期
                        if (dtOrder.Rows[i][12].ToString() == (dtOrder.Columns[j].Caption))
                        {
                            dtOrder.Rows[i][j] = s.iCount.ToString() + SPLIT + s.iBacklog.ToString()+SPLIT+"(s "+s.iSurpLus.ToString()+")";
                        }
                        else//预订日期
                        {
                            dtOrder.Rows[i][j] = s.iCount.ToString() + SPLIT + "(s "+s.iSurpLus.ToString()+")";
                        }


                        

                    }

                }
            }
        }
        private void changeViewColor()
        {
            int i, j;
            for (i = 0; i < dataGridViewP.Rows.Count; i++)
            {
                for (j = ICOLUMNS; j < dataGridViewP.Columns.Count; j++)
                {
                    string[] sT = dataGridViewP.Columns[j].Name.Split('-');

                    var alQuery = from cIme cime in alIme
                                  where (cime.sPCode == dataGridViewP.Rows[i].Cells[1].Value.ToString()) && (cime.sICode == dataGridViewP.Rows[i].Cells[6].Value.ToString()) && (cime.iYear == int.Parse(sT[0])) && (cime.iWeek == int.Parse(sT[1]))
                                  select cime;
                    if (alQuery.Count() <= 0)
                        dataGridViewP.Rows[i].Cells[j].Style.BackColor = Color.LightGray;
                    foreach (cIme s in alQuery)
                    {
                        if (s.iAcquire == 0 && s.iCount != 0) //没有区间
                        {
                            dataGridViewP.Rows[i].Cells[j].Style.BackColor = Color.LightPink;
                        }
                        else
                        {
                            if (s.iAcquire != 0)
                                dataGridViewP.Rows[i].Cells[j].Style.BackColor = Color.LightGreen;
                            else
                                dataGridViewP.Rows[i].Cells[j].Style.BackColor = Color.LightBlue;
                        }
                        break;
                    }
                }
            }
        }

        private void dataGridViewP_CellPainting(object sender, DataGridViewCellPaintingEventArgs e)
        {
            if (e.ColumnIndex < ICOLUMNS || e.RowIndex < 0)
                return;

            string[] sT = dataGridViewP.Columns[e.ColumnIndex].Name.Split('-');

            var alQuery = from cIme cime in alIme
                          where (cime.sPCode == dataGridViewP.Rows[e.RowIndex].Cells[1].Value.ToString()) && (cime.sICode == dataGridViewP.Rows[e.RowIndex].Cells[6].Value.ToString()) && (cime.iYear == int.Parse(sT[0])) && (cime.iWeek == int.Parse(sT[1]))
                            select cime;
            if (alQuery.Count() <= 0)
                dataGridViewP.Rows[e.RowIndex].Cells[e.ColumnIndex].Style.BackColor = Color.LightGray;
            foreach (cIme s in alQuery)
            {
                if (s.iAcquire == 0 && s.iCount != 0) //没有区间
                {
                    dataGridViewP.Rows[e.RowIndex].Cells[e.ColumnIndex].Style.BackColor = Color.LightPink;
                }
                else
                {
                    if (s.iAcquire != 0)
                        dataGridViewP.Rows[e.RowIndex].Cells[e.ColumnIndex].Style.BackColor = Color.LightGreen;
                    else
                        dataGridViewP.Rows[e.RowIndex].Cells[e.ColumnIndex].Style.BackColor = Color.LightBlue;
                }
                break;
            }
        

            //string[] sT = dataGridViewP.Columns[e.ColumnIndex].Name.Split('-');
            
            //var alQuery = from cIme cime in alIme
            //              where (cime.sPCode == dataGridViewP.Rows[e.RowIndex].Cells[1].Value.ToString()) && (cime.sICode == dataGridViewP.Rows[e.RowIndex].Cells[4].Value.ToString()) && (cime.iYear == int.Parse(sT[0])) && (cime.iWeek == int.Parse(sT[1]))
            //              select cime;

            //foreach (cIme s in alQuery)
            //{
            //    if (s.iAcquire==0 && s.iCount!=0 ) //没有区间
            //    {
            //        e.CellStyle.BackColor = Color.LightPink;
            //    }
            //    break;
            //}
        }

        private void dataGridViewP_CellToolTipTextNeeded(object sender, DataGridViewCellToolTipTextNeededEventArgs e)
        {
            if (e.ColumnIndex < ICOLUMNS || e.RowIndex < 0)
                return;

            string[] sT = dataGridViewP.Columns[e.ColumnIndex].Name.Split('-');

            var alQuery = from cIme cime in alIme
                          where (cime.sPCode == dataGridViewP.Rows[e.RowIndex].Cells[1].Value.ToString()) && (cime.sICode == dataGridViewP.Rows[e.RowIndex].Cells[6].Value.ToString()) && (cime.iYear == int.Parse(sT[0])) && (cime.iWeek == int.Parse(sT[1]))
                          select cime;


            foreach (cIme s in alQuery)
            {
                if (s.iAcquire == 0 && s.iCount != 0) //没有区间
                {
                    e.ToolTipText = "NO Acquired" + " Order Start:" + s.iOrderStart.ToString() + " Order End:" + s.iOrderEnd.ToString(); ;
                }
                else
                {
                    e.ToolTipText = "Start:" + s.iStart.ToString() + " End:" + s.iEnd.ToString();
                }
                break;
            }
        }

        private void btnAll_Click(object sender, EventArgs e)
        {
            initDatatable(true, true);
        }

        private void btnPCF_Click(object sender, EventArgs e)
        {
            initDatatable(true,false);
        }

        private void btnICF_Click(object sender, EventArgs e)
        {
            initDatatable(false, true);
        }

        private void dataGridViewP_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex < ICOLUMNS || e.RowIndex < 0)
                return;

            string[] sT = dataGridViewP.Columns[e.ColumnIndex].Name.Split('-');
            if (dataGridViewP.Rows[e.RowIndex].Cells[12].Value.ToString() == dataGridViewP.Columns[e.ColumnIndex].Name) //最后实际阶段没有编辑
                return;

            var alQuery = from cIme cime in alIme
                          where (cime.sPCode == dataGridViewP.Rows[e.RowIndex].Cells[1].Value.ToString()) && (cime.sICode == dataGridViewP.Rows[e.RowIndex].Cells[6].Value.ToString()) && (cime.iYear == int.Parse(sT[0])) && (cime.iWeek == int.Parse(sT[1]))
                            select cime;
            if (alQuery.Count() <= 0) //没有区间
                return;

            foreach (cIme s in alQuery)
            {
               
                FormNumberInput fni = new FormNumberInput();
                fni.numericUpDownCount.Value = s.iCount;
                fni.Location = new Point(Control.MousePosition.X,Control.MousePosition.Y);

                if (fni.ShowDialog() == DialogResult.OK)
                {
                    s.iCount = (int)fni.numericUpDownCount.Value;
                    countAmount(e.RowIndex);
                    //changeViewText();
                }
                break;
            }
            

        }



        private void printToolStripButton_Click(object sender, EventArgs e)
        {
            changeViewColor();
            string strT = "Forecast";
            PrintDGV.Print_DataGridView(dataGridViewP, strT, false, intUserLimit);
        }

        private void printPreviewToolStripButton_Click(object sender, EventArgs e)
        {
            changeViewColor();
            string strT = "Forecast";
            PrintDGV.Print_DataGridView(dataGridViewP, strT, true, intUserLimit);
        }

        private void toolStripButtonClear_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("do you want to clear data？", "Information", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) == DialogResult.No)
                return;

            sqlConn.Open();

            sqlComm.CommandText = "DELETE FROM orders";
            sqlComm.ExecuteNonQuery();
            sqlComm.CommandText = "dbcc checkident(orders,reseed,0)";
            sqlComm.ExecuteNonQuery();

            sqlConn.Close();
            MessageBox.Show("delete data finished", "Information", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            initDatatable(false, false);
        }



    }


}
