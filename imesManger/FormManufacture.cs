using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace imesManger
{
    public partial class FormManufacture : Form
    {
        private System.Data.SqlClient.SqlConnection sqlConn = new System.Data.SqlClient.SqlConnection();
        private System.Data.SqlClient.SqlCommand sqlComm = new System.Data.SqlClient.SqlCommand();
        private System.Data.SqlClient.SqlDataReader sqldr;
        private System.Data.SqlClient.SqlDataAdapter sqlDA = new System.Data.SqlClient.SqlDataAdapter();
        private System.Data.DataSet dSet = new DataSet();

        public string strConn = "";

        public int intUserLimit = 0;
        public int iRange = 1000000;

        private DataTable dtManu = new DataTable();

        public FormManufacture()
        {
            InitializeComponent();
        }

        private void FormManufacture_Load(object sender, EventArgs e)
        {
            sqlConn.ConnectionString = strConn;
            sqlComm.Connection = sqlConn;
            sqlDA.SelectCommand = sqlComm;

            this.Top = 1;
            this.Left = 1;

            dtManu.Columns.Add("pid", System.Type.GetType("System.Decimal"));//1
            dtManu.Columns.Add("Product Code", System.Type.GetType("System.String"));
            dtManu.Columns.Add("Numbers", System.Type.GetType("System.Decimal"));
            dtManu.Columns.Add("iid", System.Type.GetType("System.Decimal"));//1
            dtManu.Columns.Add("Indentor Code", System.Type.GetType("System.String"));
            dtManu.Columns.Add("TAC ID", System.Type.GetType("System.Decimal"));
            dtManu.Columns.Add("TAC Code", System.Type.GetType("System.String"));
            dtManu.Columns.Add("TAC Init Number", System.Type.GetType("System.String"));
            dtManu.Columns.Add("Last Number", System.Type.GetType("System.Decimal"));
            dtManu.Columns.Add("End Number", System.Type.GetType("System.Decimal"));
            dtManu.Columns.Add("Total Number", System.Type.GetType("System.Decimal"));
            dtManu.Columns.Add("Check", System.Type.GetType("System.Decimal"));
            dtManu.Columns.Add("backlog", System.Type.GetType("System.Decimal"));

            initDatatable(false,false);
        }
        private void setSTAUS()
        {
            toolStripStatusLabelC.Text = "Number of Factory production Rows: " + dataGridViewP.RowCount.ToString();
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
            sqlComm.CommandText = "SELECT   actual.ID, actual.[Product ID], actual.[Product Code], actual.[Indentor ID], actual.[Indentor Code], actual.[Start Number], actual.[End Number], actual.[Acquire ID], actual.Date, actual.Year, actual.[Num of Week], actual.[TAC ID], actual.[TAC Code], actual.[Total Number] FROM actual INNER JOIN (SELECT   [Product Code], [Indentor Code], MAX([Total Number]) AS [Total Number] FROM      actual AS actual_1 GROUP BY [Product Code], [Indentor Code]) AS A ON actual.[Product Code] = A.[Product Code] AND  actual.[Indentor Code] = A.[Indentor Code] AND actual.[Total Number] = A.[Total Number]";
            if (bPCode && textBoxPC.Text.Trim() != "")
            {
                sqlComm.CommandText += " WHERE (actual.[Product Code] LIKE N'%" + textBoxPC.Text.Trim() + "%')";
                bFirst = false;
            }
            if (bICode && textBoxIC.Text.Trim() != "")
            {
                if (bFirst)
                    sqlComm.CommandText += " WHERE (actual.[Indentor Code] LIKE N'%" + textBoxIC.Text.Trim() + "%')";
                else
                    sqlComm.CommandText += " AND (actual.[Indentor Code] LIKE N'%" + textBoxIC.Text.Trim() + "%')";
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

            object[] oTemp = new object[13];
            dtManu.Clear();

            int i;
            for (i = 0; i < dSet.Tables["pi"].Rows.Count; i++)
            {
                oTemp[0] = dSet.Tables["pi"].Rows[i][0];
                oTemp[1] = dSet.Tables["pi"].Rows[i][2];
                oTemp[2] = dSet.Tables["pi"].Rows[i][6];
                oTemp[3] = dSet.Tables["pi"].Rows[i][3];
                oTemp[4] = dSet.Tables["pi"].Rows[i][5];
                oTemp[5] = 0; //TAC ID
                oTemp[6] = ""; //TAC Code
                oTemp[7] = 0; //init number
                oTemp[8] = 0; //start
                oTemp[9] = 0; //end 
                oTemp[10] = 0;//total number
                oTemp[11] = 0;//check
                oTemp[12] = 0;//backlog


                var q1 = from dt1 in dSet.Tables["maxnumber"].AsEnumerable()//查询最后
                         where (dt1.Field<string>("Product Code") == oTemp[1].ToString()) && (dt1.Field<string>("Indentor Code") == oTemp[4].ToString())//条件
                         select dt1;

                if (q1.Count() > 0)
                {
                    foreach (var item in q1)//显示查询结果
                    {
                        oTemp[6] = item.Field<string>("TAC Code");
                        oTemp[8] = item.Field<int>("End Number");
                        break;
                    }
                }
                else //以前没有记录
                {
                    var q2 = from dt2 in dSet.Tables["TAC"].AsEnumerable()//查询TAC区间
                             where (dt2.Field<string>("Product Code") == oTemp[1].ToString()) && (dt2.Field<string>("Indentor Code") == oTemp[4].ToString()) && (dt2.Field<int>("Init Number") == 0)//条件
                             select dt2;

                    foreach (var item in q2)//显示查询结果
                    {
                        oTemp[5] = item.Field<int>("ID");
                        oTemp[6] = item.Field<string>("TAC Code");
                        oTemp[7] = item.Field<int>("Init Number");
                        break;
                    }

                }

                dtManu.Rows.Add(oTemp);
            }

            dataGridViewP.DataSource = dtManu;
            for (i = 0; i < dataGridViewP.ColumnCount; i++)
            {
                dataGridViewP.Columns[i].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                dataGridViewP.Columns[i].ReadOnly = true;
            }
            dataGridViewP.Columns[6].ReadOnly = false;
            dataGridViewP.Columns[9].ReadOnly = false;
            dataGridViewP.Columns[12].ReadOnly = false;

            

            dataGridViewP.Columns[0].Visible = false;
            dataGridViewP.Columns[3].Visible = false;
            dataGridViewP.Columns[5].Visible = false;
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
            initDatatable(false,false);
        }
        private void dataGridViewP_CellValidating(object sender, DataGridViewCellValidatingEventArgs e)
        {
            int intOut = 0;
            if (dataGridViewP.Rows[e.RowIndex].IsNewRow)
                return;

            ClassIm.ClearDataGridViewErrorText(dataGridViewP);
            switch (e.ColumnIndex)
            {
                case 6:
                    if (e.FormattedValue.ToString() == "") break;

                    //取得TAC
                    var q2 = from dt2 in dSet.Tables["TAC"].AsEnumerable()//查询TAC区间
                             where (dt2.Field<string>("Product Code") == dataGridViewP.Rows[e.RowIndex].Cells[1].Value.ToString()) && (dt2.Field<string>("Indentor Code") == dataGridViewP.Rows[e.RowIndex].Cells[4].Value.ToString()) && (dt2.Field<string>("TAC Code") == e.FormattedValue.ToString())//条件
                             select dt2;

                    if (q2.Count() <= 0)
                    {
                        dataGridViewP.Rows[e.RowIndex].Cells[6].ErrorText = "TAC Code error";
                        e.Cancel = true;
                    }
                    break;

                 case 9:
                    if (e.FormattedValue.ToString() == "") break;
                    if (int.TryParse(e.FormattedValue.ToString(), out intOut))
                    {
                        if (intOut < 0)
                        {
                            dataGridViewP.Rows[e.RowIndex].Cells[9].ErrorText = "data error";
                            e.Cancel = true;
                        }
                        if (intOut > iRange)
                        {
                            dataGridViewP.Rows[e.RowIndex].Cells[9].ErrorText = "data error";
                            e.Cancel = true;
                        }
                    }
                    else
                    {
                        dataGridViewP.Rows[e.RowIndex].Cells[9].ErrorText = "data error";
                        e.Cancel = true;
                    }
                    break;
                 case 12:

                    if (e.FormattedValue.ToString() == "") break;
                    if (int.TryParse(e.FormattedValue.ToString(), out intOut))
                    {
                        if (intOut < 0)
                        {
                            dataGridViewP.Rows[e.RowIndex].Cells[12].ErrorText = "data error";
                            e.Cancel = true;
                        }
                        if (intOut > iRange)
                        {
                            dataGridViewP.Rows[e.RowIndex].Cells[12].ErrorText = "data error";
                            e.Cancel = true;
                        }
                    }
                    else
                    {
                        dataGridViewP.Rows[e.RowIndex].Cells[12].ErrorText = "data error";
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

            this.dataGridViewP.CellValidating -= dataGridViewP_CellValidating;
            ClassIm.ClearDataGridViewErrorText(dataGridViewP);

            for (int i = 0; i < dtManu.Rows.Count; i++)
            {
                dtManu.Rows[i][11] = 0;
                //TAC 检验
                if (dtManu.Rows[i][6].ToString() == "")
                {
                    dtManu.Rows[i][11] = 1;
                }

                var q2 = from dt2 in dSet.Tables["TAC"].AsEnumerable()//查询TAC区间
                         where (dt2.Field<string>("Product Code") == dtManu.Rows[i][1].ToString()) && (dt2.Field<string>("Indentor Code") == dtManu.Rows[i][4].ToString()) && (dt2.Field<string>("TAC Code") == dtManu.Rows[i][6].ToString())//条件
                         select dt2;

                if (q2.Count() <= 0)
                {
                    dtManu.Rows[i][11] = 1;
                    dtManu.Rows[i][5] = 0;
                    dtManu.Rows[i][6] = "";
                    dtManu.Rows[i][7] = 0;
                }
                else
                {
                    foreach (var item in q2)//显示查询结果
                    {
                        dtManu.Rows[i][5] = item.Field<int>("ID");
                        dtManu.Rows[i][6] = item.Field<string>("TAC Code");
                        dtManu.Rows[i][7] = item.Field<int>("Init Number");
                        break;
                    }
                }

                if (int.Parse(dtManu.Rows[i][11].ToString()) != 0)
                    continue;

                //结束号检验

                if (int.Parse(dtManu.Rows[i][9].ToString()) > iRange || int.Parse(dtManu.Rows[i][9].ToString()) <= 0 || int.Parse(dtManu.Rows[i][9].ToString()) <= int.Parse(dtManu.Rows[i][8].ToString()))
                {
                    dtManu.Rows[i][11] = 2;
                }

                if (dtManu.Rows[i][12].ToString() == "")
                {
                    dtManu.Rows[i][11] = 2;
                }
                if (int.Parse(dtManu.Rows[i][12].ToString()) < 0)
                {
                    dtManu.Rows[i][12] = 2;
                }
                //取值大于跳转
                if (int.Parse(dtManu.Rows[i][12].ToString()) > iRange)
                {
                    dtManu.Rows[i][12] = 2;
                }

                if (int.Parse(dtManu.Rows[i][11].ToString()) != 0)
                    continue;

                dtManu.Rows[i][10] = int.Parse(dtManu.Rows[i][9].ToString()) + int.Parse(dtManu.Rows[i][7].ToString());

            }

            for (int i = 0; i < dataGridViewP.RowCount; i++)
            {
                switch (dataGridViewP.Rows[i].Cells[11].Value.ToString())
                {
                    case "1":
                        dataGridViewP.Rows[i].Cells[6].ErrorText = "TAC Code Error";
                        break;
                    //case "2":
                    //    dataGridViewP.Rows[i].Cells[10].ErrorText = "NO TAC Define";
                    //    break;
                    case "2":
                        dataGridViewP.Rows[i].Cells[9].ErrorText = "number error";
                        break;

                    default:
                        break;

                }

            }

            this.dataGridViewP.CellValidating += dataGridViewP_CellValidating;
            //dataGridViewP.EndEdit();

            return bCheck;
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
                    if(int.Parse(dataGridViewP.Rows[i].Cells[11].Value.ToString())!=0)
                        continue;

                    //查重
                    sqlComm.CommandText = "SELECT ID, [Product Code], [Indentor Code], Year, [Num of Week] FROM actual WHERE   ([Product ID] = N'" + dataGridViewP.Rows[i].Cells[0].Value.ToString() + "') AND ([Indentor ID] = N'" + dataGridViewP.Rows[i].Cells[3].Value.ToString() + "') AND (Year = " + dateTimePickerP.Value.Year.ToString() + ") AND ([Num of Week] = N'" + ClassIm.GetWeek(dateTimePickerP.Value).ToString() + "')";
                    sqldr=sqlComm.ExecuteReader();

                    if (!sqldr.HasRows)
                    {
                        sqlComm.CommandText = "INSERT INTO actual([Product ID], [Product Code], [Indentor ID], [Indentor Code], [Start Number], [End Number], [Acquire ID], Date,  Year, [Num of Week], [TAC ID], [TAC Code], [Total Number]) VALUES (" + dataGridViewP.Rows[i].Cells[0].Value.ToString() + ", N'" + dataGridViewP.Rows[i].Cells[1].Value.ToString() + "', " + dataGridViewP.Rows[i].Cells[3].Value.ToString() + ", N'" + dataGridViewP.Rows[i].Cells[4].Value.ToString() + "', " + dataGridViewP.Rows[i].Cells[8].Value.ToString() + ", " + dataGridViewP.Rows[i].Cells[9].Value.ToString() + ", 0, CONVERT(DATETIME, '" + System.DateTime.Now.ToShortDateString() + " 00:00:00', 102), " + dateTimePickerP.Value.Year.ToString() + ", N'" + ClassIm.GetWeek(dateTimePickerP.Value).ToString() + "', " + dataGridViewP.Rows[i].Cells[5].Value.ToString() + ", N'" + dataGridViewP.Rows[i].Cells[6].Value.ToString() + "', " + dataGridViewP.Rows[i].Cells[10].Value.ToString() + ")";
                    }
                    else
                    {
                        sqlComm.CommandText = "UPDATE actual SET [End Number] = " + dataGridViewP.Rows[i].Cells[9].Value.ToString() + ", Date = CONVERT(DATETIME, '" + System.DateTime.Now.ToShortDateString() + " 00:00:00', 102), [TAC ID] = " + dataGridViewP.Rows[i].Cells[5].Value.ToString() + ", [TAC Code] = N'" + dataGridViewP.Rows[i].Cells[6].Value.ToString() + "', [Total Number] = " + dataGridViewP.Rows[i].Cells[10].Value.ToString()+" WHERE   ([Product ID] = N'" + dataGridViewP.Rows[i].Cells[0].Value.ToString() + "') AND ([Indentor ID] = N'" + dataGridViewP.Rows[i].Cells[3].Value.ToString() + "') AND (Year = " + dateTimePickerP.Value.Year.ToString() + ") AND ([Num of Week] = N'" + ClassIm.GetWeek(dateTimePickerP.Value).ToString() + "')";;
                    }
                    sqldr.Close();
                    sqlComm.ExecuteNonQuery();

                    //backlog
                    if (dataGridViewP.Rows[i].Cells[12].Value.ToString() != "0")
                    {
                        //查重
                        sqlComm.CommandText = "SELECT   ID, [Product ID], [Product Code], [Indentor ID], [Indentor Code], [Start Number], [End Number], [Acquire ID], Date, Year, [Num of Week], Numbers, [Order Start Number], [Order End Number], backlog FROM orders WHERE ([Product ID] = N'" + dataGridViewP.Rows[i].Cells[0].Value.ToString() + "') AND ([Indentor ID] = N'" + dataGridViewP.Rows[i].Cells[3].Value.ToString() + "') AND (Year = " + dateTimePickerP.Value.Year.ToString() + ") AND ([Num of Week] = N'" + ClassIm.GetWeek(dateTimePickerP.Value).ToString() + "')";
                        sqldr = sqlComm.ExecuteReader();

                        if (!sqldr.HasRows)
                        {
                            sqlComm.CommandText = "INSERT INTO orders ([Product ID], [Product Code], [Indentor ID], [Indentor Code], [Start Number], [End Number], [Acquire ID], Date, Year,  [Num of Week], Numbers, [Order Start Number], [Order End Number], backlog) VALUES   (" + dataGridViewP.Rows[i].Cells[0].Value.ToString() + ", N'" + dataGridViewP.Rows[i].Cells[1].Value.ToString() + "', " + dataGridViewP.Rows[i].Cells[3].Value.ToString() + ", N'" + dataGridViewP.Rows[i].Cells[4].Value.ToString() + "', 0, 0, 0, CONVERT(DATETIME, '" + System.DateTime.Now.ToShortDateString() + " 00:00:00', 102), " + dateTimePickerP.Value.Year.ToString() + ", N'" + ClassIm.GetWeek(dateTimePickerP.Value).ToString() + "', 0, 0, 0, " + dataGridViewP.Rows[i].Cells[12].Value.ToString() + ")";

                        }
                        else
                        {
                            sqlComm.CommandText = "UPDATE orders SET [backlog] = " + dataGridViewP.Rows[i].Cells[12].Value.ToString() + ", Date = CONVERT(DATETIME, '" + System.DateTime.Now.ToShortDateString() + " 00:00:00', 102) WHERE ([Product ID] = N'" + dataGridViewP.Rows[i].Cells[0].Value.ToString() + "') AND ([Indentor ID] = N'" + dataGridViewP.Rows[i].Cells[3].Value.ToString() + "') AND (Year = " + dateTimePickerP.Value.Year.ToString() + ") AND ([Num of Week] = N'" + ClassIm.GetWeek(dateTimePickerP.Value).ToString() + "')"; ;
                        }
                        sqldr.Close();
                        sqlComm.ExecuteNonQuery();
                    }
                    


                }

                sqlta.Commit();
                MessageBox.Show("Manufacture finished");
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

        private void printPreviewToolStripButton_Click(object sender, EventArgs e)
        {
            string strT = "Manufacture";
            PrintDGV.Print_DataGridView(dataGridViewP, strT, true, intUserLimit);
        }

        private void printToolStripButton_Click(object sender, EventArgs e)
        {
            string strT = "Manufacture";
            PrintDGV.Print_DataGridView(dataGridViewP, strT, false, intUserLimit);
        }

        private void toolStripButtonInit_Click(object sender, EventArgs e)
        {
            int i;

            System.Data.SqlClient.SqlTransaction sqlta;
            sqlConn.Open();
            sqlta = sqlConn.BeginTransaction();
            sqlComm.Transaction = sqlta;
            try
            {
                for (i = 0; i < dataGridViewP.RowCount; i++)
                {

                    //查重
                    sqlComm.CommandText = "SELECT ID, [Product Code], [Indentor Code], Year, [Num of Week] FROM actual WHERE   ([Product ID] = N'" + dataGridViewP.Rows[i].Cells[0].Value.ToString() + "') AND ([Indentor ID] = N'" + dataGridViewP.Rows[i].Cells[3].Value.ToString() + "')";
                    sqldr = sqlComm.ExecuteReader();

                    if (!sqldr.HasRows)
                    {
                        sqlComm.CommandText = "INSERT INTO actual([Product ID], [Product Code], [Indentor ID], [Indentor Code], [Start Number], [End Number], [Acquire ID], Date,  Year, [Num of Week], [TAC ID], [TAC Code], [Total Number]) VALUES (" + dataGridViewP.Rows[i].Cells[0].Value.ToString() + ", N'" + dataGridViewP.Rows[i].Cells[1].Value.ToString() + "', " + dataGridViewP.Rows[i].Cells[3].Value.ToString() + ", N'" + dataGridViewP.Rows[i].Cells[4].Value.ToString() + "', " + dataGridViewP.Rows[i].Cells[8].Value.ToString() + ", " + dataGridViewP.Rows[i].Cells[9].Value.ToString() + ", 0, CONVERT(DATETIME, '" + System.DateTime.Now.ToShortDateString() + " 00:00:00', 102), " + dateTimePickerP.Value.Year.ToString() + ", N'" + ClassIm.GetWeek(dateTimePickerP.Value).ToString() + "',0, N'', 0)";
                    }
                    sqldr.Close();
                    sqlComm.ExecuteNonQuery();
                }

                sqlta.Commit();
                MessageBox.Show("Manufacture Initialize finished");
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
            initDatatable(false, false);
        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("do you want to clear data？", "Information", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) == DialogResult.No)
                return;

            sqlConn.Open();
            sqlComm.CommandText = "DELETE FROM actual";
            sqlComm.ExecuteNonQuery();
            sqlComm.CommandText = "dbcc checkident(actual,reseed,0)";
            sqlComm.ExecuteNonQuery();
            sqlConn.Close();
            MessageBox.Show("delete data finished", "Information", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            initDatatable(false, false);
        }
    }
}
