using System;
using System.Data;
using System.Linq;
using System.Windows.Forms;

namespace imesManger
{
    public partial class FormAcquire : Form
    {
        private System.Data.SqlClient.SqlConnection sqlConn = new System.Data.SqlClient.SqlConnection();
        private System.Data.SqlClient.SqlCommand sqlComm = new System.Data.SqlClient.SqlCommand();
        private System.Data.SqlClient.SqlDataReader sqldr;
        private System.Data.SqlClient.SqlDataAdapter sqlDA = new System.Data.SqlClient.SqlDataAdapter();
        private System.Data.DataSet dSet = new DataSet();

        public string strConn = "";

        public int intUserLimit = 0;
        private DataTable dtAcquire = new DataTable();

        public int iRange = 1000000;

        public FormAcquire()
        {
            InitializeComponent();
        }

        private void FormAcquire_Load(object sender, EventArgs e)
        {
            sqlConn.ConnectionString = strConn;
            sqlComm.Connection = sqlConn;
            sqlDA.SelectCommand = sqlComm;

            this.Top = 1;
            this.Left = 1;

            dtAcquire.Columns.Add("pid", System.Type.GetType("System.Decimal"));//1
            dtAcquire.Columns.Add("Product Code", System.Type.GetType("System.String"));
            dtAcquire.Columns.Add("Numbers", System.Type.GetType("System.Decimal"));
            dtAcquire.Columns.Add("iid", System.Type.GetType("System.Decimal"));//1
            dtAcquire.Columns.Add("Indentor Code", System.Type.GetType("System.String"));
            dtAcquire.Columns.Add("Initial Number", System.Type.GetType("System.Decimal"));
            dtAcquire.Columns.Add("Start Number", System.Type.GetType("System.Decimal"));
            dtAcquire.Columns.Add("End Number", System.Type.GetType("System.Decimal"));
            dtAcquire.Columns.Add("TAC ID", System.Type.GetType("System.Decimal"));
            dtAcquire.Columns.Add("TAC Code", System.Type.GetType("System.String"));
            dtAcquire.Columns.Add("Quantity", System.Type.GetType("System.Decimal"));
            dtAcquire.Columns.Add("Check", System.Type.GetType("System.Decimal"));
            

            initDatatable(false,false);
        }

        private void setSTAUS()
        {
            toolStripStatusLabelC.Text = "Number of Acquires : " + dataGridViewP.RowCount.ToString();
        }

        private void initDatatable(bool bPCode, bool bICode)
        {
            bool bFirst = true;

            sqlConn.Open();
            bFirst = true;
            sqlComm.CommandText = "SELECT product.ID, product.[Product Name], product.[Product Code], indentor.ID AS iID, indentor.[Indentor Name], indentor.[Indentor Code], product.[Number of IMEI] FROM product CROSS JOIN indentor";
            if (bPCode && textBoxPC.Text.Trim() != "")
            {
                sqlComm.CommandText += " WHERE (product.[Product Code] LIKE N'%" + textBoxPC.Text.Trim() + "%')";
                bFirst = false;
            }
            if (bICode && textBoxIC.Text.Trim() != "")
            {
                if (bFirst)
                    sqlComm.CommandText += " WHERE (indentor.[Indentor Code] LIKE N'%" + textBoxIC.Text.Trim() + "%')";
                else
                    sqlComm.CommandText += " AND (indentor.[Indentor Code] LIKE N'%" + textBoxIC.Text.Trim() + "%')";
            }
            if (dSet.Tables.Contains("pi")) dSet.Tables["pi"].Clear();
            sqlDA.Fill(dSet, "pi");

            bFirst = true;
            sqlComm.CommandText = "SELECT [Product Code], [Indentor Code], MAX([End Number]) AS [MAX NUMBER] FROM acquire GROUP BY [Product Code], [Indentor Code]";
            if (bPCode && textBoxPC.Text.Trim() != "")
            {
                sqlComm.CommandText += " HAVING    ([Product Code] LIKE N'%" + textBoxPC.Text.Trim() + "%')";
                bFirst = false;
            }
            if (bICode && textBoxIC.Text.Trim() != "")
            {
                if (bFirst)
                    sqlComm.CommandText += " HAVING ([Indentor Code] LIKE N'%" + textBoxIC.Text.Trim() + "%')";
                else
                    sqlComm.CommandText += " AND ([Indentor Code] LIKE N'%" + textBoxIC.Text.Trim() + "%')";
            }
            if (dSet.Tables.Contains("maxnumber")) dSet.Tables["maxnumber"].Clear();
            sqlDA.Fill(dSet, "maxnumber");

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
            if (dSet.Tables.Contains("TAC")) dSet.Tables["TAC"].Clear();
            sqlDA.Fill(dSet, "TAC");
            sqlComm.CommandText += " ORDER BY [Init Number]";

            sqlConn.Close();

            object[] oTemp = new object[12];
            dtAcquire.Clear();

            int i;
            for (i = 0; i < dSet.Tables["pi"].Rows.Count; i++)
            {
                oTemp[0] = dSet.Tables["pi"].Rows[i][0];
                oTemp[1] = dSet.Tables["pi"].Rows[i][2];
                oTemp[2] = dSet.Tables["pi"].Rows[i][6];
                oTemp[3] = dSet.Tables["pi"].Rows[i][3];
                oTemp[4] = dSet.Tables["pi"].Rows[i][5];
                oTemp[5] = 0; //initial number
                oTemp[6] = 0; //start
                oTemp[7] = 0; //end 
                oTemp[8] = 0; //TAC ID
                oTemp[9] = ""; //TAC Code
                oTemp[10] = 0; //Quantity
                oTemp[11] = 0;

                var q1 = from dt1 in dSet.Tables["maxnumber"].AsEnumerable()//查询最后
                         where (dt1.Field<string>("Product Code") == oTemp[1].ToString()) && (dt1.Field<string>("Indentor Code") == oTemp[4].ToString())//条件
                         select dt1;

                foreach (var item in q1)//显示查询结果
                {
                    oTemp[5] = item.Field<int>("MAX NUMBER")+1;
                    break;
                }
                oTemp[6] = oTemp[5]; oTemp[7] = oTemp[6];

                var q2 = from dt2 in dSet.Tables["TAC"].AsEnumerable()//查询TAC区间
                         where (dt2.Field<string>("Product Code") == oTemp[1].ToString()) && (dt2.Field<string>("Indentor Code") == oTemp[4].ToString()) && (dt2.Field<int>("Init Number") >= int.Parse(oTemp[5].ToString()) && (dt2.Field<int>("Init Number") <= int.Parse(oTemp[5].ToString())+iRange))//条件
                         select dt2;
                foreach (var item in q2)//显示查询结果
                {
                    oTemp[8] = item.Field<int>("ID");
                    oTemp[9] = item.Field<string>("TAC Code");
                    break;
                }
                dtAcquire.Rows.Add(oTemp);
            }

            dataGridViewP.DataSource = dtAcquire;
            for (i = 0; i < dataGridViewP.ColumnCount - 2; i++)
            {
                dataGridViewP.Columns[i].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                dataGridViewP.Columns[i].ReadOnly = true;
            }
            dataGridViewP.Columns[i].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;

            dataGridViewP.Columns[0].Visible = false;
            dataGridViewP.Columns[3].Visible = false;
            dataGridViewP.Columns[8].Visible = false;
            dataGridViewP.Columns[11].Visible = false;
            setSTAUS();

        }

        private void btnPCF_Click(object sender, EventArgs e)
        {
            initDatatable(true, false);
        }

        private void btnICF_Click(object sender, EventArgs e)
        {
            initDatatable(false, true);
        }

        private void btnAll_Click(object sender, EventArgs e)
        {
            initDatatable(false, false);
        }

        private void toolStripButtonEDIT_Click(object sender, EventArgs e)
        {
            int i;

            if (!countAmount())
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
                
                for (i = 0; i < dataGridViewP.RowCount; i++)
                {
                    if (int.Parse(dataGridViewP.Rows[i].Cells[11].Value.ToString()) != 0)
                        continue;

                    if (int.Parse(dataGridViewP.Rows[i].Cells[10].Value.ToString()) <= 0)
                        continue;
                    //dateTimePickerP.Value.w

                    sqlComm.CommandText = "INSERT INTO acquire([Product ID], [Product Code], [Indentor ID], [Indentor Code], [Start Number], [End Number], Date, Year,  [Num of Week], [TAC ID], [TAC Code]) VALUES   (" + dataGridViewP.Rows[i].Cells[0].Value.ToString() + ", N'" + dataGridViewP.Rows[i].Cells[1].Value.ToString() + "', " + dataGridViewP.Rows[i].Cells[3].Value.ToString() + ", N'" + dataGridViewP.Rows[i].Cells[4].Value.ToString() + "', " + dataGridViewP.Rows[i].Cells[6].Value.ToString() + ", " + dataGridViewP.Rows[i].Cells[7].Value.ToString() + ", '" + System.DateTime.Now.ToShortDateString() + " 00:00:00', " + dateTimePickerP.Value.Year.ToString() + ",N'" + ClassIm.GetWeek(dateTimePickerP.Value).ToString() + "', " + dataGridViewP.Rows[i].Cells[8].Value.ToString() + ", N'" + dataGridViewP.Rows[i].Cells[9].Value.ToString() + "')"; 


                    
                    sqlComm.ExecuteNonQuery();
                }

                sqlta.Commit();
                MessageBox.Show("Acquire finished");
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
            initDatatable(false,false);

        }

        private void dataGridViewP_CellValidating(object sender, DataGridViewCellValidatingEventArgs e)
        {
            if (dataGridViewP.Rows[e.RowIndex].IsNewRow)
                return;
            int intOut = 0;

            ClassIm.ClearDataGridViewErrorText(dataGridViewP);
            switch (e.ColumnIndex)
            {
                case 10:

                    if (e.FormattedValue.ToString() == "") break;
                    if (int.TryParse(e.FormattedValue.ToString(), out intOut))
                    {
                        if (intOut < 0)
                        {
                            dataGridViewP.Rows[e.RowIndex].Cells[10].ErrorText = "data error";
                            e.Cancel = true;
                        }
                    }
                    else
                    {
                        dataGridViewP.Rows[e.RowIndex].Cells[10].ErrorText = "data error";
                        e.Cancel = true;
                    }
                    break;
                default:
                    break;
            }
            //dataGridViewP.EndEdit();

        }

        private void dataGridViewP_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            MessageBox.Show("dataformat error", "information", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        private void dataGridViewP_RowValidating(object sender, DataGridViewCellCancelEventArgs e)
        {
            countAmount();
        }

        private bool countAmount()
        {
            bool bCheck = true;
            int itemp = 0, itemp1 = 0;
            int iR = 0;
            bool bFirst=true;

            this.dataGridViewP.CellValidating -= dataGridViewP_CellValidating;
            ClassIm.ClearDataGridViewErrorText(dataGridViewP);

            for (int i = 0; i < dtAcquire.Rows.Count; i++)
            {
                dtAcquire.Rows[i][11] = 0;
                if (dtAcquire.Rows[i][10].ToString() == "")
                {
                    dtAcquire.Rows[i][11]=1;
                }
                if (int.Parse(dtAcquire.Rows[i][10].ToString()) < 0 )
                {
                    dtAcquire.Rows[i][11] = 1;
                }

                //取值大于跳转
                if (int.Parse(dtAcquire.Rows[i][10].ToString()) > iRange)
                {
                    dtAcquire.Rows[i][11] = 3;
                }

                

                if(int.Parse(dtAcquire.Rows[i][11].ToString())!=0)
                    continue;

                dtAcquire.Rows[i][11] = 0;

                //判断区间
                itemp = int.Parse(dtAcquire.Rows[i][5].ToString());
                itemp1 = itemp + int.Parse(dtAcquire.Rows[i][10].ToString()) * int.Parse(dtAcquire.Rows[i][2].ToString())-1;
                if (!checkBoxCon.Checked) //不连续计算
                {
                    iR = itemp / iRange;
                    iR = iR * iRange;
                    bFirst = true;
                    while (true)
                    {
                        if (itemp >= iR && itemp1 <= iR + iRange) //找到区间
                        {
                            dtAcquire.Rows[i][11] = 0;
                            dtAcquire.Rows[i][6] = itemp;
                            dtAcquire.Rows[i][7] = itemp1;

                            dtAcquire.Rows[i][8] = 0;
                            dtAcquire.Rows[i][9] = "";


                            var q2 = from dt2 in dSet.Tables["TAC"].AsEnumerable()//查询TAC区间
                                     where (dt2.Field<string>("Product Code") == dtAcquire.Rows[i][1].ToString()) && (dt2.Field<string>("Indentor Code") == dtAcquire.Rows[i][4].ToString()) && (dt2.Field<int>("Init Number") == iR)//条件
                                     select dt2;

                            foreach (var item in q2)//显示查询结果
                            {
                                dtAcquire.Rows[i][8] = item.Field<int>("ID");
                                dtAcquire.Rows[i][9] = item.Field<string>("TAC Code");
                                break;
                            }
                            break;

                        }
                        else
                        {
                            iR += iRange;
                            itemp = iR;
                            itemp1 = itemp + int.Parse(dtAcquire.Rows[i][10].ToString()) * int.Parse(dtAcquire.Rows[i][2].ToString());
                        }
                    }
                }
                else //连续计算
                {
                    dtAcquire.Rows[i][11] = 0;
                    dtAcquire.Rows[i][6] = itemp;
                    dtAcquire.Rows[i][7] = itemp1;

                    dtAcquire.Rows[i][8] = 0;
                    dtAcquire.Rows[i][9] = "";


                    var q2 = from dt2 in dSet.Tables["TAC"].AsEnumerable()//查询TAC区间
                             where (dt2.Field<string>("Product Code") == dtAcquire.Rows[i][1].ToString()) && (dt2.Field<string>("Indentor Code") == dtAcquire.Rows[i][4].ToString()) && (dt2.Field<int>("Init Number") == iR)//条件
                             select dt2;

                    foreach (var item in q2)//显示查询结果
                    {
                        dtAcquire.Rows[i][8] = item.Field<int>("ID");
                        dtAcquire.Rows[i][9] = item.Field<string>("TAC Code");
                        break;
                    }
                }


            }

            for (int i = 0; i < dataGridViewP.RowCount; i++)
            {
                switch(dataGridViewP.Rows[i].Cells[11].Value.ToString())
                {
                    case "1":
                        dataGridViewP.Rows[i].Cells[10].ErrorText = "please input number";
                        break;
                    //case "2":
                    //    dataGridViewP.Rows[i].Cells[10].ErrorText = "NO TAC Define";
                    //    break;
                    case "3":
                        dataGridViewP.Rows[i].Cells[10].ErrorText = "OUT OF TAC Define";
                        break;

                    default:
                        break;

                }
                
            }

            this.dataGridViewP.CellValidating += dataGridViewP_CellValidating;
            //dataGridViewP.EndEdit();

            return bCheck;
        }

        private void printPreviewToolStripButton_Click(object sender, EventArgs e)
        {
            string strT = "Acquire";
            PrintDGV.Print_DataGridView(dataGridViewP, strT, true, intUserLimit);
        }

        private void printToolStripButton_Click(object sender, EventArgs e)
        {
            string strT = "Acquire";
            PrintDGV.Print_DataGridView(dataGridViewP, strT, false, intUserLimit);
        }

        private void toolStripButtonClear_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("do you want to clear data？", "Information", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) == DialogResult.No)
                return;

            sqlConn.Open();

            sqlComm.CommandText = "DELETE FROM acquire";
            sqlComm.ExecuteNonQuery();
            sqlComm.CommandText = "dbcc checkident(acquire,reseed,0)";
            sqlComm.ExecuteNonQuery();

            sqlConn.Close();
            MessageBox.Show("delete data finished", "Information", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            initDatatable(false, false);
        }




    }
}
