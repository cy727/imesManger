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
    public partial class FormManageOffline : Form
    {
        //private System.Data.SqlClient.SqlConnection sqlConn = new System.Data.SqlClient.SqlConnection();
        //private System.Data.SqlClient.SqlCommand sqlComm = new System.Data.SqlClient.SqlCommand();
        //private System.Data.SqlClient.SqlDataReader sqldr;
        //private System.Data.SqlClient.SqlDataAdapter sqlDA = new System.Data.SqlClient.SqlDataAdapter();
        private System.Data.DataSet dSet = new DataSet();

        //public string strConn = "";

        public int intUserLimit = 0;
        private DataTable dtOrderFore = new DataTable();
        private DataTable dtOrderFore1 = new DataTable();

        public System.Data.DataTable dtProduct;
        public System.Data.DataTable dtCustumor;
        public System.Data.DataTable dtTAC;
        public System.Data.DataTable dtAcquire;
        public System.Data.DataTable dtActual;
        public System.Data.DataTable dtOrder;
        public System.Data.DataTable dtBacklog;

        private System.Data.DataTable dtPI = new DataTable();

        private System.Data.DataTable dtPI1;
        private System.Data.DataTable dtTAC1;
        private System.Data.DataTable dtAcquire1;
        private System.Data.DataTable dtActual1;
        private System.Data.DataTable dtOrder1;
        //private System.Data.DataTable dtBacklog1;

        public int iNumofWeek = 11;
        public int iNumofPass = 1;

        private ArrayList alIme = new ArrayList();
        public int iRange = 1000000;

        private const int ICOLUMNS = 21;
        private const char SPLIT = ':';
        private const int IWEEKSAFTER = 2; //缺省几周后预测

        public FormManageOffline()
        {
            InitializeComponent();
        }

        private void FormManageOffline_Load(object sender, EventArgs e)
        {

            this.Top = 1;
            this.Left = 1;

            comboBoxIC.DataSource = dtCustumor;
            comboBoxIC.DisplayMember = "Customer Code";
            comboBoxIC.Text = "";

            comboBoxPC.DataSource = dtProduct;
            comboBoxPC.DisplayMember = "Product Code";
            comboBoxPC.Text = "";

            dateTimePickerBefore.Value = System.DateTime.Now.AddDays(IWEEKSAFTER * 7);

          
            dtPI.Columns.Add("ID", System.Type.GetType("System.Decimal"));//0
            dtPI.Columns.Add("Product Name", System.Type.GetType("System.String"));
            dtPI.Columns.Add("Product Code", System.Type.GetType("System.String"));
            dtPI.Columns.Add("iID", System.Type.GetType("System.Decimal"));
            dtPI.Columns.Add("Customer Name", System.Type.GetType("System.String"));
            dtPI.Columns.Add("Customer Code", System.Type.GetType("System.String"));
            dtPI.Columns.Add("Number of IMEI", System.Type.GetType("System.Decimal"));
            dtPI.Columns.Add("Failure Rate", System.Type.GetType("System.Decimal"));
            dtPI.Columns.Add("Factory", System.Type.GetType("System.String"));

            object[] oTemp = new object[9];

            for (int i = 0; i < dtCustumor.Rows.Count; i++)
            {
                for (int j = 0; j < dtProduct.Rows.Count; j++)
                {
                    oTemp[0] = dtProduct.Rows[j][0];
                    oTemp[1] = dtProduct.Rows[j][2];
                    oTemp[2] = dtProduct.Rows[j][1];
                    oTemp[3] = dtCustumor.Rows[i][0];
                    oTemp[4] = dtCustumor.Rows[i][2];
                    oTemp[5] = dtCustumor.Rows[i][1];
                    oTemp[6] = dtProduct.Rows[j][3];
                    oTemp[7] = dtProduct.Rows[j][4];
                    oTemp[8] = dtProduct.Rows[j][5];
                    dtPI.Rows.Add(oTemp);
                }
            }

            initDatatable(false, false);

        }

        private void setSTAUS()
        {
            toolStripStatusLabelC.Text = "Number of ORDER Rows: " + dataGridViewP.RowCount.ToString();
        }

        private void initDatatable(bool bPCode, bool bICode)
        {
            int i, j, k;
            int iTemp = 0, iTemp1 = 0;
            string sTemp = "";
            bool bFirst = true;

            dataGridViewP.RowValidating -= dataGridViewP_RowValidating;
            dataGridViewP.CellValidating -= dataGridViewP_CellValidating;

            dtOrderFore.Columns.Clear();
            dtOrderFore.Columns.Add("pid", System.Type.GetType("System.Decimal"));//0
            dtOrderFore.Columns.Add("Product Code", System.Type.GetType("System.String"));//1
            dtOrderFore.Columns.Add("Product Name", System.Type.GetType("System.String"));//2
            dtOrderFore.Columns.Add("SIM Slot", System.Type.GetType("System.Decimal"));//3
            dtOrderFore.Columns.Add("Factory", System.Type.GetType("System.String"));//4
            dtOrderFore.Columns.Add("Failure Rate", System.Type.GetType("System.Decimal"));//5
            dtOrderFore.Columns.Add("iid", System.Type.GetType("System.Decimal"));//6
            dtOrderFore.Columns.Add("Customer Code", System.Type.GetType("System.String"));//7
            dtOrderFore.Columns.Add("Customer Name", System.Type.GetType("System.String"));//8
            dtOrderFore.Columns.Add("TAC ID", System.Type.GetType("System.Decimal"));//9
            dtOrderFore.Columns.Add("TAC Code", System.Type.GetType("System.String"));//10
            dtOrderFore.Columns.Add("TAC Init", System.Type.GetType("System.Decimal"));//11
            dtOrderFore.Columns.Add("Manufacture Number", System.Type.GetType("System.Decimal"));//12
            dtOrderFore.Columns.Add("Used IMEI Quantity", System.Type.GetType("System.Decimal"));//13
            dtOrderFore.Columns.Add("Production week", System.Type.GetType("System.String"));//14
            dtOrderFore.Columns.Add("Request TAC Code", System.Type.GetType("System.String"));//15
            dtOrderFore.Columns.Add("Request TAC Init", System.Type.GetType("System.Decimal"));//16
            dtOrderFore.Columns.Add("Request Number", System.Type.GetType("System.Decimal"));//17
            dtOrderFore.Columns.Add("Request Date", System.Type.GetType("System.String"));//18
            dtOrderFore.Columns.Add("Remaining", System.Type.GetType("System.Decimal")); //19总剩余 
            dtOrderFore.Columns.Add("Section Remaining", System.Type.GetType("System.Decimal")); //20区间剩余


            for (i = iNumofPass; i > 0; i--)
            {
                sTemp = ClassIm.GetWeek(dateTimePickerP.Value, -1 * i);
                sTemp = ClassIm.sYear + "-" + sTemp;

                dtOrderFore.Columns.Add(sTemp, System.Type.GetType("System.String"));
            }

            for (i = 0; i < iNumofWeek; i++)
            {
                sTemp = ClassIm.GetWeek(dateTimePickerP.Value, i);
                sTemp = ClassIm.sYear + "-" + sTemp;

                dtOrderFore.Columns.Add(sTemp, System.Type.GetType("System.String"));
            }

            #region 生成dataset
            //生成PI
            dtPI1 = dtPI.Clone();dtPI1.TableName="pi";

            //空白不判断
            if (comboBoxPC.Text.Trim() == "")
                bPCode = false;
            if (comboBoxIC.Text.Trim() == "")
                bICode = false;
            
            foreach (DataRow dr in dtPI.Rows)
            {

                if (bPCode && bICode) //全部过滤
                {
                    if (dr["Product Code"].ToString().Contains(comboBoxPC.Text.Trim()) && dr["Customer Code"].ToString().Contains(comboBoxIC.Text.Trim()))
                    {
                        dtPI1.ImportRow(dr);
                    }
                }
                if (!bPCode && !bICode) //全不过滤
                {
                    dtPI1.ImportRow(dr);
                }

                if (bPCode && !bICode) //过滤product
                {
                    if (dr["Product Code"].ToString().Contains(comboBoxPC.Text.Trim()))
                    {
                        dtPI1.ImportRow(dr);
                    }
                }
                if (!bPCode && bICode) //过滤costumer
                {
                    if (dr["Customer Code"].ToString().Contains(comboBoxIC.Text.Trim()))
                    {
                        dtPI1.ImportRow(dr);
                    }
                }
            }
            if (dSet.Tables.Contains("pi")) dSet.Tables.Remove("pi");
            dSet.Tables.Add(dtPI1);  //条件过滤后加入dset

            //生成TAC
            dtTAC1 = dtTAC.Clone(); dtTAC1.TableName = "TAC";
            foreach (DataRow dr in dtTAC.Rows)
            {
                if (bPCode && bICode) //
                {
                    if (dr["Product Code"].ToString().Contains(comboBoxPC.Text.Trim()) && dr["Customer Code"].ToString().Contains(comboBoxIC.Text.Trim()))
                    {
                        dtTAC1.ImportRow(dr);
                    }
                }
                if (!bPCode && !bICode) //
                {
                    dtTAC1.ImportRow(dr);
                }
                if (bPCode && !bICode) //
                {
                    if (dr["Product Code"].ToString().Contains(comboBoxPC.Text.Trim()))
                    {
                        dtTAC1.ImportRow(dr);
                    }
                }
                if (!bPCode && bICode) //
                {
                    if (dr["Customer Code"].ToString().Contains(comboBoxIC.Text.Trim()))
                    {
                        dtTAC1.ImportRow(dr);
                    }
                }
            }
            if (dSet.Tables.Contains("TAC")) dSet.Tables.Remove("TAC");
            dSet.Tables.Add(dtTAC1);

            //acquire
            dtAcquire1 = dtAcquire.Clone(); dtAcquire1.TableName = "acquire";
            foreach (DataRow dr in dtAcquire.Rows)
            {
                if (bPCode && bICode) //
                {
                    if (dr["Product Code"].ToString().Contains(comboBoxPC.Text.Trim()) && dr["Customer Code"].ToString().Contains(comboBoxIC.Text.Trim()))
                    {
                        dtAcquire1.ImportRow(dr);
                    }
                }
                if (!bPCode && !bICode) //
                {
                    dtAcquire1.ImportRow(dr);
                }
                if (bPCode && !bICode) //
                {
                    if (dr["Product Code"].ToString().Contains(comboBoxPC.Text.Trim()))
                    {
                        dtAcquire1.ImportRow(dr);
                    }
                }
                if (!bPCode && bICode) //
                {
                    if (dr["Customer Code"].ToString().Contains(comboBoxIC.Text.Trim()))
                    {
                        dtAcquire1.ImportRow(dr);
                    }
                }
            }
            if (dSet.Tables.Contains("acquire")) dSet.Tables.Remove("acquire");
            dSet.Tables.Add(dtAcquire1);

            //actual
            dtActual1 = dtActual.Clone(); dtActual1.TableName = "actual";
            foreach (DataRow dr in dtActual.Rows)
            {
                if (bPCode && bICode) //
                {
                    if (dr["Product Code"].ToString().Contains(comboBoxPC.Text.Trim()) && dr["Customer Code"].ToString().Contains(comboBoxIC.Text.Trim()))
                    {
                        dtActual1.ImportRow(dr);
                    }
                }
                if (!bPCode && !bICode) //
                {
                    dtActual1.ImportRow(dr);
                }
                if (bPCode && !bICode) //
                {
                    if (dr["Product Code"].ToString().Contains(comboBoxPC.Text.Trim()))
                    {
                        dtActual1.ImportRow(dr);
                    }
                }
                if (!bPCode && bICode) //
                {
                    if (dr["Customer Code"].ToString().Contains(comboBoxIC.Text.Trim()))
                    {
                        dtActual1.ImportRow(dr);
                    }
                }
            }
            if (dSet.Tables.Contains("actual")) dSet.Tables.Remove("actual");
            dSet.Tables.Add(dtActual1);


            //order
            dtOrder1 = dtOrder.Clone(); dtOrder1.TableName = "orders";
            foreach (DataRow dr in dtOrder.Rows)
            {
                if (bPCode && bICode) //
                {
                    if (dr["Product Code"].ToString().Contains(comboBoxPC.Text.Trim()) && dr["Customer Code"].ToString().Contains(comboBoxIC.Text.Trim()))
                    {
                        dtOrder1.ImportRow(dr);
                    }
                }
                if (!bPCode && !bICode) //
                {
                    dtOrder1.ImportRow(dr);
                }
                if (bPCode && !bICode) //
                {
                    if (dr["Product Code"].ToString().Contains(comboBoxPC.Text.Trim()))
                    {
                        dtOrder1.ImportRow(dr);
                    }
                }
                if (!bPCode && bICode) //
                {
                    if (dr["Customer Code"].ToString().Contains(comboBoxIC.Text.Trim()))
                    {
                        dtOrder1.ImportRow(dr);
                    }
                }
            }
            if (dSet.Tables.Contains("orders")) dSet.Tables.Remove("orders");
            dSet.Tables.Add(dtOrder1);

            #endregion


            object[] oTemp = new object[ICOLUMNS + iNumofWeek + iNumofPass];
            dtOrderFore.Clear();


            for (i = 0; i < dSet.Tables["pi"].Rows.Count; i++)
            {
                oTemp[0] = dSet.Tables["pi"].Rows[i][0];
                oTemp[1] = dSet.Tables["pi"].Rows[i][2];
                oTemp[2] = dSet.Tables["pi"].Rows[i][1];
                oTemp[3] = dSet.Tables["pi"].Rows[i][6];
                oTemp[4] = dSet.Tables["pi"].Rows[i][8];
                oTemp[5] = dSet.Tables["pi"].Rows[i][7];
                oTemp[6] = dSet.Tables["pi"].Rows[i][3];
                oTemp[7] = dSet.Tables["pi"].Rows[i][5];
                oTemp[8] = dSet.Tables["pi"].Rows[i][4];

                //TAC Manufacture
                oTemp[9] = 0;
                oTemp[10] = "";
                oTemp[11] = 0;
                oTemp[12] = 0;
                oTemp[13] = 0;
                oTemp[14] = "2000-1";

                var q2 = from dt1 in dSet.Tables["actual"].AsEnumerable()//查询最后 生产
                         where (dt1.Field<string>("Product Code") == oTemp[1].ToString()) && (dt1.Field<string>("Customer Code") == oTemp[7].ToString())//条件
                         select dt1;

                foreach (var item in q2)//显示查询结果
                {
                    //TAC
                    oTemp[9] = item.Field<decimal>("TAC ID");
                    oTemp[10] = item.Field<string>("TAC Code").ToString();

                    var qTAC = from dt2 in dSet.Tables["TAC"].AsEnumerable()//得到TAC初始值
                               where (dt2.Field<string>("Product Code") == oTemp[1].ToString()) && (dt2.Field<string>("Customer Code") == oTemp[7].ToString()) && (dt2.Field<string>("TAC Code") == oTemp[10].ToString())//条件
                               select dt2;
                    foreach (var item1 in qTAC)//显示查询结果
                    {
                        oTemp[11] = item1.Field<decimal>("Init Number");
                        break;
                    }

                    //Manufacture
                    oTemp[12] = item.Field<decimal>("End Number");
                    oTemp[13] = item.Field<decimal>("Used IMEI Quantity");
                    oTemp[14] = item.Field<decimal>("Year").ToString() + "-" + item.Field<string>("Num of Week");
                    break;
                }

                //acquire request
                oTemp[15] = "";
                oTemp[16] = 0;
                oTemp[17] = 0;
                oTemp[18] = "";
                var q3 = from dt1 in dSet.Tables["acquire"].AsEnumerable()//查询最后
                         where (dt1.Field<string>("Product Code") == oTemp[1].ToString()) && (dt1.Field<string>("Customer Code") == oTemp[7].ToString())//条件
                         orderby (dt1.Field<decimal>("End Number") + dt1.Field<decimal>("Init Number")) descending
                         select dt1;

                foreach (var item in q3)//显示查询结果
                {
                    oTemp[15] = item.Field<string>("TAC Code");
                    oTemp[16] = item.Field<decimal>("Init Number");
                    oTemp[17] = item.Field<decimal>("End Number");
                    oTemp[18] = item.Field<decimal>("Year").ToString() + "-" + item.Field<string>("Num of Week");
                    break;
                }

                //剩余号码
                iTemp = 0; iTemp1 = 0;
                getSurplusNow(oTemp[1].ToString(), oTemp[7].ToString(), int.Parse(oTemp[11].ToString()), int.Parse(oTemp[12].ToString()),out iTemp,out iTemp1);
                oTemp[19] = iTemp;
                oTemp[20] = iTemp1;

                //ORDER numbers
                k = ICOLUMNS;
                for (j = iNumofPass; j > 0; j--)
                {
                    oTemp[k] = 0;
                    //得到日期
                    string[] sT = dtOrderFore.Columns[k].Caption.Split('-');
                    if (sT.Length == 2)
                    {
                        var q4 = from dt1 in dSet.Tables["orders"].AsEnumerable()//查询最后
                                 where (dt1.Field<string>("Product Code") == oTemp[1].ToString()) && (dt1.Field<string>("Customer Code") == oTemp[7].ToString()) && (dt1.Field<decimal>("Year").ToString() == decimal.Parse(sT[0]).ToString()) && (dt1.Field<string>("Num of Week") == sT[1])//条件
                                 select dt1;
                        foreach (var item in q4)//显示查询结果
                        {
                            oTemp[k] = item.Field<decimal>("Numbers");
                            break;
                        }
                    }
                    k++;
                }

                for (j = 0; j < iNumofWeek; j++)
                {
                    oTemp[k] = 0;
                    //得到日期
                    string[] sT = dtOrderFore.Columns[k].Caption.Split('-');
                    if (sT.Length == 2)
                    {
                        var q5 = from dt1 in dSet.Tables["orders"].AsEnumerable()//查询最后
                                 where (dt1.Field<string>("Product Code") == oTemp[1].ToString()) && (dt1.Field<string>("Customer Code") == oTemp[7].ToString()) && (dt1.Field<decimal>("Year").ToString() == decimal.Parse(sT[0]).ToString()) && (dt1.Field<string>("Num of Week") == sT[1])//条件
                                 select dt1;
                        foreach (var item in q5)//显示查询结果
                        {
                            oTemp[k] = item.Field<decimal>("Numbers");
                            break;
                        }
                    }
                    k++;
                }

                if (checkBoxTAC.Checked)  //屏蔽所有没有TAC值的记录
                {
                    if (oTemp[10].ToString() == "" || oTemp[15].ToString() == "")
                        continue;
                }




                dtOrderFore.Rows.Add(oTemp);
            }

            dataGridViewP.DataSource = dtOrderFore;
            for (i = 0; i < dataGridViewP.Columns.Count; i++)
            {
                dataGridViewP.Columns[i].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                dataGridViewP.Columns[i].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            }
            for (i = 0; i < dataGridViewP.Columns.Count; i++)
            {
                dataGridViewP.Columns[i].ReadOnly = true;
            }


            dataGridViewP.Columns[1].DefaultCellStyle.BackColor = Color.LightBlue;
            dataGridViewP.Columns[8].DefaultCellStyle.BackColor = Color.LightBlue;
            dataGridViewP.Columns[8].Frozen = true;

            //***********************
            dataGridViewP.Columns[0].Visible = false;
            dataGridViewP.Columns[6].Visible = false;
            dataGridViewP.Columns[9].Visible = false;
            dataGridViewP.Columns[11].Visible = false;
            dataGridViewP.Columns[12].Visible = false;
            dataGridViewP.Columns[16].Visible = false;
            //***********/
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
                string[] sT = dataGridViewP.Rows[i].Cells[14].Value.ToString().Split('-'); // 生产日期
                iTemp = int.Parse(sT[0]) * 100 + int.Parse(sT[1]);
                var q10 = from dt1 in dSet.Tables["orders"].AsEnumerable()//大于起始时间Order
                          where (dt1.Field<string>("Product Code") == dataGridViewP.Rows[i].Cells[1].Value.ToString()) && (dt1.Field<string>("Customer Code") == dataGridViewP.Rows[i].Cells[7].Value.ToString()) && (dt1.Field<decimal>("Year") * 100 + int.Parse(dt1.Field<string>("Num of Week")) >= iTemp)
                          orderby (dt1.Field<decimal>("Year") * 100 + int.Parse(dt1.Field<string>("Num of Week")))
                          select dt1;
                foreach (var item in q10)//查询结果
                {
                    cIme cime = new cIme(int.Parse(dataGridViewP.Rows[i].Cells[0].Value.ToString()), dataGridViewP.Rows[i].Cells[1].Value.ToString(), int.Parse(dataGridViewP.Rows[i].Cells[6].Value.ToString()), dataGridViewP.Rows[i].Cells[7].Value.ToString(), int.Parse(dataGridViewP.Rows[i].Cells[3].Value.ToString()));
                    cime.iYear = (int)(item.Field<decimal>("Year"));
                    cime.iWeek = int.Parse(item.Field<string>("Num of Week"));

                    cime.iCount =(int)(item.Field<decimal>("Numbers"));
                    //cime.iBacklog = item.Field<decimal>("backlog");


                    //backlog 补充
                    cime.iBacklog = 0;
                    var q11 = from dt1 in dtBacklog.AsEnumerable()//大于起始时间Order
                          where (dt1.Field<string>("Product Code") == dataGridViewP.Rows[i].Cells[1].Value.ToString()) && (dt1.Field<string>("Customer Code") == dataGridViewP.Rows[i].Cells[7].Value.ToString()) && (dt1.Field<decimal>("Year")==cime.iYear && (dt1.Field<string>("Num of Week")==cime.iWeek.ToString()))
                          select dt1;
                    foreach (var item1 in q11)//显示查询结果
                    {
                        cime.iBacklog = (int)(item1.Field<decimal>("backlog"));
                    }

                    cime.dFailureRate = decimal.Parse(dataGridViewP.Rows[i].Cells[5].Value.ToString());
                    alIme.Add(cime);
                }

                //加入未有数据
                for (j = ICOLUMNS; j < dataGridViewP.Columns.Count; j++)
                {
                    //判断是有效日期
                    string[] sT1 = dataGridViewP.Columns[j].Name.Split('-');
                    iTemp1 = int.Parse(sT1[0]) * 100 + int.Parse(sT1[1]);
                    if (iTemp1 < iTemp) continue;

                    var alQuery = from cIme cime in alIme
                                  where (cime.sPCode == dataGridViewP.Rows[i].Cells[1].Value.ToString()) && (cime.sICode == dataGridViewP.Rows[i].Cells[7].Value.ToString()) && (cime.iYear == int.Parse(sT1[0])) && (cime.iWeek == int.Parse(sT1[1]))
                                  select cime;

                    //foreach (cIme s in alQuery)
                    //{
                    //}
                    if (alQuery.Count() < 1) //数据库没有记录，加入预订为0的数据
                    {
                        cIme cime1 = new cIme(int.Parse(dataGridViewP.Rows[i].Cells[0].Value.ToString()), dataGridViewP.Rows[i].Cells[1].Value.ToString(), int.Parse(dataGridViewP.Rows[i].Cells[6].Value.ToString()), dataGridViewP.Rows[i].Cells[7].Value.ToString(), int.Parse(dataGridViewP.Rows[i].Cells[3].Value.ToString()));
                        cime1.iYear = int.Parse(sT1[0]);
                        cime1.iWeek = int.Parse(sT1[1]);

                        cime1.iCount = 0;
                        cime1.dFailureRate = decimal.Parse(dataGridViewP.Rows[i].Cells[5].Value.ToString());
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
        private void getSurplusNow(string sPCode, string sICode, int iInitNumber, int iManNumber, out int iSurplusAll, out int iSurplusNow)
        {
            if (sPCode.Trim() == "" || sICode.Trim() == "")
            {
                iSurplusAll = 0; iSurplusNow = 0;
                return;
            }

            int i;
            iSurplusAll = 0; iSurplusNow = 0;
            var qAcquire = from dt1 in dSet.Tables["acquire"].AsEnumerable()//查询最后
                           where (dt1.Field<string>("Product Code") == sPCode) && (dt1.Field<string>("Customer Code") == sICode)//条件
                           orderby dt1.Field<decimal>("End Number") + dt1.Field<decimal>("Init Number") ascending
                           select dt1;

            foreach (var item in qAcquire)//显示查询结果
            {
                if (iManNumber + iInitNumber > item.Field<decimal>("End Number") + item.Field<decimal>("Init Number")) //当前大于此区间
                    continue;

                if (iManNumber + iInitNumber < item.Field<decimal>("Start Number") + item.Field<decimal>("Init Number")) //当前小于此区间
                {
                    iSurplusAll += (int)(item.Field<decimal>("End Number") - item.Field<decimal>("Start Number"));
                    continue;
                }
                if (iManNumber + iInitNumber >= item.Field<decimal>("Start Number") + item.Field<decimal>("Init Number") && iManNumber + iInitNumber <= item.Field<decimal>("End Number") + item.Field<decimal>("Init Number")) //在此区间内
                {
                    iSurplusNow = (int)(item.Field<decimal>("End Number")) - iManNumber;
                    iSurplusAll += iSurplusNow;
                    continue;
                }

            }

            return ;
        }


        private int getSurplus(string sPCode, string sICode, int iManNumber)
        {
            if (sPCode.Trim() == "" || sICode == "")
                return 0;

            int i;
            int iSurplus = 0;
            var qAcquire = from dt1 in dSet.Tables["acquire"].AsEnumerable()//查询最后
                           where (dt1.Field<string>("Product Code") == sPCode) && (dt1.Field<string>("Customer Code") == sICode)//条件
                           orderby dt1.Field<decimal>("End Number") ascending
                           select dt1;

            foreach (var item in qAcquire)//显示查询结果
            {
                if (iManNumber > item.Field<decimal>("End Number"))
                    continue;

                if (iManNumber < item.Field<decimal>("Start Number"))
                {
                    iSurplus += (int)(item.Field<decimal>("End Number") - item.Field<decimal>("Start Number"));
                    continue;
                }
                if (iManNumber >= item.Field<decimal>("Start Number") && iManNumber <= item.Field<decimal>("End Number"))
                {
                    iSurplus += (int)(item.Field<decimal>("End Number")) - iManNumber;
                    continue;
                }

            }

            return iSurplus;
        }

        /*
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
        */

        private void btnRe_Click(object sender, EventArgs e)
        {
            initDatatable(false, false);
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
            int iStartYear = 0, iStartWeek = 0;
            int iEndYear = 0, iEndWeek = 0;
            int iStartNo = 0, iEndNo = 0;

            int iRStart = 0, iRend = 0;

            int iAcquireNO = 0;
            int iTemp = 0;
            bool bTrue = true;

            int iSurplus = 0;

            //manageArrayList(); //重置用户输入
            this.dataGridViewP.CellValidating -= dataGridViewP_CellValidating;
            ClassIm.ClearDataGridViewErrorText(dataGridViewP);

            //得到计算行
            if (iRows < 0)
            {
                iRStart = 0;
                iRend = dataGridViewP.Rows.Count;
            }
            else
            {
                iRStart = iRows;
                iRend = iRows + 1;
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
                string[] sT = dataGridViewP.Rows[i].Cells[14].Value.ToString().Split('-');
                iStartYear = int.Parse(sT[0]); iStartWeek = int.Parse(sT[1]);
                iStartNo = int.Parse(dataGridViewP.Rows[i].Cells[12].Value.ToString());
                if (iStartNo != 0)
                    iStartNo++;//起始号
                iStartNo += int.Parse(dataGridViewP.Rows[i].Cells[11].Value.ToString());//TAC区间


                //本周计算,考虑backlog
                bool bNext = false; //判断是否调区间

                iEndWeek = iStartWeek;
                iEndYear = iStartYear;
                var alQuery1 = from cIme cime in alIme
                               where (cime.sPCode == dataGridViewP.Rows[i].Cells[1].Value.ToString()) && (cime.sICode == dataGridViewP.Rows[i].Cells[7].Value.ToString()) && (cime.iYear == iEndYear) && (cime.iWeek == iEndWeek)
                               select cime;
                foreach (cIme s in alQuery1)
                {
                    iAcquireNO = 0;
                    bTrue = true;
                    while (bTrue)
                    {
                        //找到开始号区间
                        var q3 = from dt1 in dSet.Tables["acquire"].AsEnumerable()//查询最后
                                 where (dt1.Field<string>("Product Code") == s.sPCode) && (dt1.Field<string>("Customer Code") == s.sICode) && (dt1.Field<decimal>("End Number") + dt1.Field<decimal>("Init Number") >= iStartNo)//条件
                                 orderby dt1.Field<decimal>("End Number")
                                 select dt1;

                        if (q3.Count() < 1)  //没有区间，所有区间已不符合，跳出
                        {
                            s.iAcquire = 0;
                            break;
                        }
                        bFirst = true; s.iAcquire = 0;
                        foreach (var item in q3) //已经找到区间
                        {
                            if (!bFirst) //第二次循环，以此次申请的第一位数为起始数
                            {
                                s.bNext = true;
                                bNext = true;
                                iStartNo = (int)(item.Field<decimal>("Start Number")) + (int)(item.Field<decimal>("Init Number"));
                            }
                            iEndNo = iStartNo + (int)Math.Ceiling((double)s.iBacklog * (double)s.iNum * ((double)s.dFailureRate / 100.0 + 1.0)) - 1;
                            //s.iCount = 0;
                            if (iEndNo <= item.Field<decimal>("End Number") + item.Field<decimal>("Init Number")) //找到可以容纳的区间
                            {
                                s.iAcquire = (int)(item.Field<decimal>("ID"));
                                s.iAcStart = (int)(item.Field<decimal>("Start Number"));
                                s.iAcEnd = (int)(item.Field<decimal>("End Number"));

                                s.sTac = item.Field<string>("TAC Code");
                                if (s.sTac != dataGridViewP.Rows[i].Cells[10].Value.ToString())
                                    s.bNextTAC = true;

                                if (bNext)
                                    s.bNext = true;

                                s.iStart = iStartNo-(int)(item.Field<decimal>("Init Number")); s.iEnd = iEndNo-(int)(item.Field<decimal>("Init Number"));

                                iStartNo = iEndNo + 1;
                                //s.iSurpLus = getSurplus(s.sPCode, s.sICode, iEndNo);
                                getSurplusNow(s.sPCode, s.sICode, (int)(item.Field<decimal>("Init Number")),s.iEnd,out s.iSurpLus,out s.iSurpLusNow);
                                bTrue = false;
                                break;
                            }
                            else //区间不够，跳到下一区间
                            {
                                //得到下一个申请区间
                                bFirst = false;
                            }
                            

                        }
                        break;
                    } //while
                    //得到预计计数
                    if (s.iAcquire == 0)
                    {
                        s.iOrderStart = iStartNo;
                        s.iOrderEnd = iStartNo + (int)Math.Ceiling((double)s.iBacklog * (double)s.iNum * ((double)s.dFailureRate / 100.0 + 1.0)) - 1;
                        iSurplus = -1 * (int)Math.Ceiling((double)s.iBacklog * (double)s.iNum * ((double)s.dFailureRate / 100.0 + 1.0));
                        s.iSurpLus = iSurplus;
                        s.iSurpLusNow = iSurplus;
                        iStartNo = s.iOrderEnd + 1;
                    }
                    break;
                }//s

                //计算以后周
                k = 1;
                

                while (true)
                {
                    //从制造下一周开始
                    iEndWeek = int.Parse(ClassIm.GetWeek(iStartYear, iStartWeek, k));
                    iEndYear = int.Parse(ClassIm.sYear);
                    k++;
                    var alQuery = from cIme cime in alIme
                                  where (cime.sPCode == dataGridViewP.Rows[i].Cells[1].Value.ToString()) && (cime.sICode == dataGridViewP.Rows[i].Cells[7].Value.ToString()) && (cime.iYear == iEndYear) && (cime.iWeek == iEndWeek)
                                  select cime;

                    if (alQuery.Count() < 1) //没有记录
                        break;

                    foreach (cIme s in alQuery)
                    {
                        iAcquireNO = 0;

                        bTrue = true;
                        bFirst = true; s.iAcquire = 0;
                        while (bTrue)
                        {

                            //找到开始号区间
                            var q3 = from dt1 in dSet.Tables["acquire"].AsEnumerable()//查询最后
                                     where (dt1.Field<string>("Product Code") == s.sPCode) && (dt1.Field<string>("Customer Code") == s.sICode) &&  (dt1.Field<decimal>("End Number") + dt1.Field<decimal>("Init Number") >= iStartNo)//条件
                                     orderby dt1.Field<decimal>("End Number") + dt1.Field<decimal>("Init Number")
                                     select dt1;


                            if (q3.Count() < 1)  //没有区间
                            {
                                s.iAcquire = 0;
                                break;
                            }
                            
                            foreach (var item in q3)
                            {
                                if (!bFirst) //第二次循环，以此次申请的第一位数为起始数
                                {
                                    s.bNext = true;
                                    bNext = true;
                                    iStartNo = (int)(item.Field<decimal>("Start Number")) + (int)(item.Field<decimal>("Init Number"));
                                }
                                iEndNo = iStartNo + (int)Math.Ceiling((double)s.iCount * (double)s.iNum * ((double)s.dFailureRate / 100.0 + 1.0)) - 1;


                                if (iEndNo <= item.Field<decimal>("End Number") + item.Field<decimal>("Init Number")) //找到区间
                                {
                                    s.iAcquire = (int)(item.Field<decimal>("ID"));
                                    s.iAcStart = (int)(item.Field<decimal>("Start Number"));
                                    s.iAcEnd = (int)(item.Field<decimal>("End Number"));

                                    s.sTac = item.Field<string>("TAC Code");
                                    if (s.sTac != dataGridViewP.Rows[i].Cells[10].Value.ToString())
                                        s.bNextTAC = true;

                                    s.iStart = iStartNo - (int)item.Field<decimal>("Init Number"); s.iEnd = iEndNo - (int)item.Field<decimal>("Init Number");
                                    iStartNo = iEndNo + 1;
                                    if (bNext)
                                        s.bNext = true;

                                    //s.iSurpLus = getSurplus(s.sPCode, s.sICode, iEndNo);
                                    getSurplusNow(s.sPCode, s.sICode, (int)(item.Field<decimal>("Init Number")), s.iEnd, out s.iSurpLus, out s.iSurpLusNow);

                                    bTrue = false;
                                    break;
                                }
                                else //区间不够，跳到下一区间
                                {
                                    bFirst = false;

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
                            s.iSurpLusNow = -1 * ((int)Math.Ceiling((double)s.iCount * (double)s.iNum * ((double)s.dFailureRate / 100.0 + 1.0)));
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

            for (i = 0; i < dtOrderFore.Rows.Count; i++)
            {
                for (j = ICOLUMNS; j < dtOrderFore.Columns.Count; j++)
                {
                    string[] sT = dtOrderFore.Columns[j].ColumnName.Split('-');

                    var alQuery = from cIme cime in alIme
                                  where (cime.sPCode == dtOrderFore.Rows[i][1].ToString()) && (cime.sICode == dtOrderFore.Rows[i][7].ToString()) && (cime.iYear == int.Parse(sT[0])) && (cime.iWeek == int.Parse(sT[1]))
                                  select cime;
                    if (alQuery.Count() <= 0)
                        continue;
                    foreach (cIme s in alQuery)
                    {
                        //最终日期
                        if (dtOrderFore.Rows[i][14].ToString() == (dtOrderFore.Columns[j].Caption))
                        {
                            dtOrderFore.Rows[i][j] = s.iCount.ToString() + SPLIT + s.iBacklog.ToString() + SPLIT + "(s " + s.iSurpLusNow.ToString() + "/"+s.iSurpLus.ToString()+")";
                        }
                        else//预订日期
                        {
                            dtOrderFore.Rows[i][j] = s.iCount.ToString() + SPLIT + "(s " + s.iSurpLusNow.ToString() + "/"+s.iSurpLus.ToString()+")";
                        }




                    }

                }
            }
        }
        
        private void changeViewColor()
        {
            int i, j,iTemp=0;
            for (i = 0; i < dataGridViewP.Rows.Count; i++)
            {
                dataGridViewP.Rows[i].Cells[1].Style.BackColor = Color.LightBlue;
                dataGridViewP.Rows[i].Cells[8].Style.BackColor = Color.LightBlue;

                for (j = ICOLUMNS; j < dataGridViewP.Columns.Count; j++)
                {
                    string[] sT = dataGridViewP.Columns[j].Name.Split('-');

                    var alQuery = from cIme cime in alIme
                                  where (cime.sPCode == dataGridViewP.Rows[i].Cells[1].Value.ToString()) && (cime.sICode == dataGridViewP.Rows[i].Cells[7].Value.ToString()) && (cime.iYear == int.Parse(sT[0])) && (cime.iWeek == int.Parse(sT[1]))
                                  select cime;
                    if (alQuery.Count() <= 0) //没有记录
                        dataGridViewP.Rows[i].Cells[j].Style.BackColor = Color.LightGray;

                    foreach (cIme s in alQuery)
                    {
                        if (j == ICOLUMNS + iNumofPass)
                            iTemp = s.iBacklog;
                        else
                            iTemp = s.iCount;

                        if (s.iAcquire == 0 && iTemp != 0) //没有区间
                        {
                            dataGridViewP.Rows[i].Cells[j].Style.BackColor = Color.Pink;
                        }
                        else
                        {
                            if (s.iAcquire != 0) //有区间
                            {
                                if (!s.bNext) //原有区间
                                    dataGridViewP.Rows[i].Cells[j].Style.BackColor = Color.LightGreen;
                                else //已经调转了区间
                                    if (s.bNextTAC)
                                        dataGridViewP.Rows[i].Cells[j].Style.BackColor = Color.Orange;
                                    else
                                        dataGridViewP.Rows[i].Cells[j].Style.BackColor = Color.LightYellow; 
                                    
                            }
                            else //没有order
                                dataGridViewP.Rows[i].Cells[j].Style.BackColor = Color.LightGray;
                        }
                        break;
                    }

                }
            }
        }
        
        private void dataGridViewP_CellPainting(object sender, DataGridViewCellPaintingEventArgs e)
        {
            int iTemp=0;
            if (e.ColumnIndex < ICOLUMNS || e.RowIndex < 0)
                return;

            string[] sT = dataGridViewP.Columns[e.ColumnIndex].Name.Split('-');

            var alQuery = from cIme cime in alIme
                          where (cime.sPCode == dataGridViewP.Rows[e.RowIndex].Cells[1].Value.ToString()) && (cime.sICode == dataGridViewP.Rows[e.RowIndex].Cells[7].Value.ToString()) && (cime.iYear == int.Parse(sT[0])) && (cime.iWeek == int.Parse(sT[1]))
                          select cime;
            if (alQuery.Count() <= 0) //没有记录
                dataGridViewP.Rows[e.RowIndex].Cells[e.ColumnIndex].Style.BackColor = Color.LightGray;

            foreach (cIme s in alQuery)
            {
                if (e.ColumnIndex == ICOLUMNS + iNumofPass)
                    iTemp = s.iBacklog;
                else
                    iTemp = s.iCount;

                if (s.iAcquire == 0 && iTemp != 0) //没有区间
                {
                    dataGridViewP.Rows[e.RowIndex].Cells[e.ColumnIndex].Style.BackColor = Color.Pink;
                }
                else
                {
                    if (s.iAcquire != 0) //有区间
                    {
                        if(!s.bNext) //原有区间
                            dataGridViewP.Rows[e.RowIndex].Cells[e.ColumnIndex].Style.BackColor = Color.LightGreen;
                        else //已经调转了区间
                            if(s.bNextTAC)
                                dataGridViewP.Rows[e.RowIndex].Cells[e.ColumnIndex].Style.BackColor = Color.Orange;
                            else
                                dataGridViewP.Rows[e.RowIndex].Cells[e.ColumnIndex].Style.BackColor = Color.LightYellow;
                    }
                    else //没有order
                        dataGridViewP.Rows[e.RowIndex].Cells[e.ColumnIndex].Style.BackColor = Color.LightGray;
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
                          where (cime.sPCode == dataGridViewP.Rows[e.RowIndex].Cells[1].Value.ToString()) && (cime.sICode == dataGridViewP.Rows[e.RowIndex].Cells[7].Value.ToString()) && (cime.iYear == int.Parse(sT[0])) && (cime.iWeek == int.Parse(sT[1]))
                          select cime;


            foreach (cIme s in alQuery)
            {
                if (s.iAcquire == 0 && s.iCount != 0) //没有区间
                {
                    e.ToolTipText = "NO Acquired" + " Order Start:" + s.iOrderStart.ToString() + " Order End:" + s.iOrderEnd.ToString(); ;
                }
                else
                {
                    e.ToolTipText = "TAC:"+s.sTac+" Start:" + s.iStart.ToString() + " End:" + s.iEnd.ToString();
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
            initDatatable(true, false);
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
                fni.Location = new Point(Control.MousePosition.X, Control.MousePosition.Y);

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

        private void btndisplayfilter_Click(object sender, EventArgs e)
        {
            //
            int iColumn = dataGridViewP.Columns.Count;
            int i, j, iYear,iWeek,iTemp;
            string sTemp="";
            bool bShow = true;

            dtOrderFore1 = dtOrderFore.Copy();
            if(checkBoxBefore.Checked) //区间
            {
                if(dateTimePickerBefore.Value<System.DateTime.Now)
                {
                    MessageBox.Show("Date wrong");
                    return;
                }

                iWeek=ClassIm.GetWeek(dateTimePickerBefore.Value);
                iYear=int.Parse( ClassIm.sYear);

                sTemp = iYear.ToString()+ "-" + iWeek.ToString();

                for (j = ICOLUMNS; j < dtOrderFore1.Columns.Count; j++)
                {
                    if(sTemp==dataGridViewP.Columns[j].Name)
                    {
                        iColumn=j;
                        break;
                    }
                }
            }

            for (i = dtOrderFore1.Rows.Count-1; i >=0; i--)
            {
                bShow = false;
                for (j = ICOLUMNS + iNumofPass; j < iColumn; j++)
                {
                    string[] sT = dataGridViewP.Columns[j].Name.Split('-');

                    var alQuery = from cIme cime in alIme
                                  where (cime.sPCode == dtOrderFore.Rows[i][1].ToString()) && (cime.sICode == dtOrderFore.Rows[i][7].ToString()) && (cime.iYear == int.Parse(sT[0])) && (cime.iWeek == int.Parse(sT[1]))
                                  select cime;

                    //if (alQuery.Count() <= 0) //没有记录
                    //    dataGridViewP.Rows[i].Cells[j].Style.BackColor = Color.LightGray;

                    foreach (cIme s in alQuery)
                    {
                        if (j == ICOLUMNS + iNumofPass)
                        {
                            if (s.iAcquire == 0 && s.iBacklog != 0) //没有区间
                            {
                                bShow = true;
                                break;
                            }

                        }
                        else
                        {
                            if (s.iAcquire == 0 && s.iCount != 0) //没有区间
                            {
                                bShow = true;
                                break;
                            }
                        }
                    }
                    if (bShow)
                        break;
                }
                if(!bShow)
                    dtOrderFore1.Rows[i].Delete();
            }

            dtOrderFore1.AcceptChanges();
            dataGridViewP.DataSource = null;
            dataGridViewP.DataSource = dtOrderFore1;

            dataGridViewP.Columns[1].DefaultCellStyle.BackColor = Color.LightBlue;
            dataGridViewP.Columns[8].DefaultCellStyle.BackColor = Color.LightBlue;
            dataGridViewP.Columns[8].Frozen = true;

            dataGridViewP.Columns[0].Visible = false;
            dataGridViewP.Columns[6].Visible = false;
            dataGridViewP.Columns[9].Visible = false;
            dataGridViewP.Columns[11].Visible = false;
            dataGridViewP.Columns[12].Visible = false;
            dataGridViewP.Columns[16].Visible = false;

            setSTAUS();
        
        }

        private void btndisplayfiltersection_Click(object sender, EventArgs e)
        {
            //
            int iColumn = dataGridViewP.Columns.Count;
            int i, j, iYear, iWeek, iTemp;
            string sTemp = "";
            bool bShow = true;

            dtOrderFore1 = dtOrderFore.Copy();
            if (checkBoxBefore.Checked) //区间
            {
                if (dateTimePickerBefore.Value < System.DateTime.Now)
                {
                    MessageBox.Show("Date wrong");
                    return;
                }

                iWeek = ClassIm.GetWeek(dateTimePickerBefore.Value);
                iYear = int.Parse(ClassIm.sYear);

                sTemp = iYear.ToString() + "-" + iWeek.ToString();

                for (j = ICOLUMNS; j < dtOrderFore1.Columns.Count; j++)
                {
                    if (sTemp == dataGridViewP.Columns[j].Name)
                    {
                        iColumn = j;
                        break;
                    }
                }
            }

            for (i = dtOrderFore1.Rows.Count - 1; i >= 0; i--)
            {
                bShow = false;
                for (j = ICOLUMNS + iNumofPass; j < iColumn; j++)
                {
                    string[] sT = dataGridViewP.Columns[j].Name.Split('-');

                    var alQuery = from cIme cime in alIme
                                  where (cime.sPCode == dtOrderFore.Rows[i][1].ToString()) && (cime.sICode == dtOrderFore.Rows[i][7].ToString()) && (cime.iYear == int.Parse(sT[0])) && (cime.iWeek == int.Parse(sT[1]))
                                  select cime;

                    //if (alQuery.Count() <= 0) //没有记录
                    //    dataGridViewP.Rows[i].Cells[j].Style.BackColor = Color.LightGray;

                    foreach (cIme s in alQuery)
                    {
                        if (j == ICOLUMNS + iNumofPass)
                        {
                            if (s.iAcquire == 0 && s.iBacklog != 0) //没有区间
                            {
                                bShow = true;
                                break;
                            }

                        }
                        else
                        {
                            if (s.iAcquire == 0 && s.iCount != 0) //没有区间
                            {
                                bShow = true;
                                break;
                            }
                        }

                        if (s.bNext)
                        {
                            bShow = true;
                            break;
                        }
                    }
                    if (bShow)
                        break;
                }
                if (!bShow)
                    dtOrderFore1.Rows[i].Delete();
            }

            dtOrderFore1.AcceptChanges();
            dataGridViewP.DataSource = null;
            dataGridViewP.DataSource = dtOrderFore1;
            setSTAUS();


            dataGridViewP.Columns[1].DefaultCellStyle.BackColor = Color.LightBlue;
            dataGridViewP.Columns[8].DefaultCellStyle.BackColor = Color.LightBlue;
            dataGridViewP.Columns[8].Frozen = true;

            dataGridViewP.Columns[0].Visible = false;
            dataGridViewP.Columns[6].Visible = false;
            dataGridViewP.Columns[9].Visible = false;
            dataGridViewP.Columns[11].Visible = false;
            dataGridViewP.Columns[12].Visible = false;
            dataGridViewP.Columns[16].Visible = false;
        }

       



    }

    class cIme
    {

        public string sPCode = "";
        public string sICode = "";
        public int iPCode = 0;
        public int iICode = 0;
        public int iNum = 0; //sim

        public int iYear = 0;
        public int iWeek = 0;

        public int iStart = 0;
        public int iEnd = 0;
        public int iCount = 0; //数量

        public int iAcquire = 0;
        public int iAcStart = 0;
        public int iAcEnd = 0;

        public int iOrderStart = 0;
        public int iOrderEnd = 0;

        public int iBacklog = 0;
        public int iSurpLus = 0;//剩余
        public int iSurpLusNow = 0;//本区间剩余

        public int iTac = 0;
        public string sTac = "";
        public int iTacInit = 0;

        public decimal dFailureRate = 0;
        public bool bNext = false;
        public bool bNextTAC = false;

        public cIme(int iPCode, string sPCode, int iICode, string sICode, int iNum) //带参数的构造函数
        {
            this.sPCode = sPCode;
            this.sICode = sICode;
            this.iNum = iNum;

            this.iPCode = iPCode;
            this.iICode = iICode;
        }
    }

}
