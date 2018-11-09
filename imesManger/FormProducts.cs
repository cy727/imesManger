using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Drawing.Printing;
using System.Data.OleDb;
using Microsoft.Office.Interop.Excel;
using System.Collections;


namespace imesManger
{
    public partial class FormProducts : Form
    {

        private System.Data.SqlClient.SqlConnection sqlConn = new System.Data.SqlClient.SqlConnection();
        private System.Data.SqlClient.SqlCommand sqlComm = new System.Data.SqlClient.SqlCommand();
        private System.Data.SqlClient.SqlDataReader sqldr;
        private System.Data.SqlClient.SqlDataAdapter sqlDA = new System.Data.SqlClient.SqlDataAdapter();
        private System.Data.DataSet dSet = new DataSet();

        public string strConn = "";

        public int intUserLimit = 0;


        private DataView dvSelect;

        public FormProducts()
        {
            InitializeComponent();
        }

        private void FormProducts_Load(object sender, EventArgs e)
        {
            sqlConn.ConnectionString = strConn;
            sqlComm.Connection = sqlConn;
            sqlDA.SelectCommand = sqlComm;

            this.Top = 1;
            this.Left = 1;

            sqlConn.Open();

            //product datatable
            sqlComm.CommandText = "SELECT ID, [Product Name], [Product Code], [Number of IMEI], [Failure Rate] FROM product ORDER BY [Product Name]";

            if (dSet.Tables.Contains("product")) dSet.Tables["product"].Clear();
            sqlDA.Fill(dSet, "product");

            sqlComm.CommandText = "SELECT ID, [Product Name], [Product Code], [Number of IMEI], [Failure Rate] FROM product ORDER BY [Product Name]";

            if (dSet.Tables.Contains("product1")) dSet.Tables.Remove("product1");
            sqlDA.Fill(dSet, "product1");

            dvSelect = new DataView(dSet.Tables["product"]);
            dataGridViewProduct.DataSource = dvSelect;
            dataGridViewProduct.Columns[0].Visible = false;
            //dataGridViewProduct.Columns[4].Visible = false;

            setSTAUS();
            sqlConn.Close();

            //initDataView();

        }

        private void setSTAUS()
        {
            toolStripStatusLabelC.Text="Number of products:"+dataGridViewProduct.RowCount.ToString();
        }

        private void initDataView(int iSel)
        {
            
            //初始化列表
            sqlConn.Open();

            sqlComm.CommandText = "SELECT ID, [Product Name], [Product Code], [Number of IMEI], [Failure Rate] FROM product ORDER BY [Product Name]";

            if (dSet.Tables.Contains("product")) dSet.Tables["product"].Clear();
            sqlDA.Fill(dSet, "product");
            sqlConn.Close();

            setSTAUS();

            if (dataGridViewProduct.Rows.Count < 1)
                return;
            if (iSel != 0)
            {
                dataGridViewProduct.Rows[0].Selected = false;
                int iRow = -1;

                for (int i = 0; i < dataGridViewProduct.Rows.Count; i++)
                {
                   if (dataGridViewProduct.Rows[i].Cells[0].Value.ToString() == iSel.ToString())
                   {
                       iRow = i;
                       break;
                    }
                }


                if (iRow != -1)
                {
                    dataGridViewProduct.Rows[iRow].Selected = true;
                    dataGridViewProduct.FirstDisplayedScrollingRowIndex = iRow;
                }


            }

        }


        private void ToolStripButtonADD_Click(object sender, EventArgs e)
        {


            dSet.Tables["product1"].Clear();
            System.Data.DataTable dt = dSet.Tables["product1"];

            FormProduct_CARD formProduct_CARD = new FormProduct_CARD();
            formProduct_CARD.strConn = strConn;
            formProduct_CARD.dt = dt;
            formProduct_CARD.iStyle = 0;

            formProduct_CARD.ShowDialog();
            initDataView(formProduct_CARD.iSelect);
        }

        private void toolStripButtonEDIT_Click(object sender, EventArgs e)
        {
            object[] oT = new object[dataGridViewProduct.ColumnCount];

            if (dataGridViewProduct.SelectedRows.Count < 1)
            {
                MessageBox.Show("please select products", "Infomation", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            dSet.Tables["product1"].Clear();
            System.Data.DataTable dt = dSet.Tables["product1"];

            for (int i = dataGridViewProduct.SelectedRows.Count-1; i >=0; i--)
            {
                for(int j=0;j<oT.Length;j++)
                    oT[j]= dataGridViewProduct.SelectedRows[i].Cells[j].Value;
                dt.Rows.Add(oT);
            }

            FormProduct_CARD formProduct_CARD = new FormProduct_CARD();
            formProduct_CARD.strConn = strConn;
            formProduct_CARD.dt = dt;
            formProduct_CARD.iStyle = 1;

            formProduct_CARD.ShowDialog();
            initDataView(formProduct_CARD.iSelect);

        }

        private void toolStripButtonDEL_Click(object sender, EventArgs e)
        {
            if (dataGridViewProduct.SelectedRows.Count < 1)
            {
                MessageBox.Show("please select products", "Infomation", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            
            if (MessageBox.Show("do you want to delete product？", "Information", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) == DialogResult.No)
                return;

            int i;
            System.Data.SqlClient.SqlTransaction sqlta;

            sqlConn.Open();
            sqlta = sqlConn.BeginTransaction();
            sqlComm.Transaction = sqlta;
            try
            {

                for (i = 0; i < dataGridViewProduct.SelectedRows.Count; i++)
                {
                    if (dataGridViewProduct.Rows[i].IsNewRow)
                    continue;

                    sqlComm.CommandText = "DELETE FROM product WHERE (ID = " + dataGridViewProduct.SelectedRows[i].Cells[0].Value.ToString() + ")";
                    sqlComm.ExecuteNonQuery();

                    sqlComm.CommandText = "DELETE FROM acquire WHERE ([Product ID] = " + dataGridViewProduct.SelectedRows[i].Cells[0].Value.ToString() + ")";
                    sqlComm.ExecuteNonQuery();

                    sqlComm.CommandText = "DELETE FROM actual WHERE ([Product ID] = " + dataGridViewProduct.SelectedRows[i].Cells[0].Value.ToString() + ")";
                    sqlComm.ExecuteNonQuery();

                    sqlComm.CommandText = "DELETE FROM buyer WHERE ([Product ID] = " + dataGridViewProduct.SelectedRows[i].Cells[0].Value.ToString() + ")";
                    sqlComm.ExecuteNonQuery();

                    sqlComm.CommandText = "DELETE FROM orders WHERE ([Product ID] = " + dataGridViewProduct.SelectedRows[i].Cells[0].Value.ToString() + ")";
                    sqlComm.ExecuteNonQuery();

                    sqlComm.CommandText = "DELETE FROM TAC WHERE ([Product ID] = " + dataGridViewProduct.SelectedRows[i].Cells[0].Value.ToString() + ")";
                    sqlComm.ExecuteNonQuery();
                 }

                    sqlta.Commit();
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
              MessageBox.Show("delete finished", "infomation", MessageBoxButtons.OK, MessageBoxIcon.Information);
              initDataView(0);
 
        }

        private void dataGridViewDJMX_DoubleClick(object sender, EventArgs e)
        {
            toolStripButtonEDIT_Click(null, null);
        }

        private void printPreviewToolStripButton_Click(object sender, EventArgs e)
        {

            string strT = "product";
            PrintDGV.Print_DataGridView(dataGridViewProduct, strT, true,intUserLimit);
        }

        private void printToolStripButton_Click(object sender, EventArgs e)
        {
            string strT = "product";
            PrintDGV.Print_DataGridView(dataGridViewProduct, strT, false,intUserLimit);
        }

        private void btnAll_Click(object sender, EventArgs e)
        {
            dvSelect.RowFilter = "";
            dvSelect.RowStateFilter = DataViewRowState.CurrentRows;
        }

        private void btnSearch_Click(object sender, EventArgs e)
        {
            if (textBoxMC.Text.Trim() == "")
                return;

            dvSelect.RowStateFilter = DataViewRowState.CurrentRows;
            dvSelect.RowFilter = " [Product Name] LIKE '%" + textBoxMC.Text.Trim() + "%'";
        }

        private void btnLocation_Click(object sender, EventArgs e)
        {
            if (textBoxMC.Text.Trim() == "")
                return;

            int iRow = -1;
            string sTemp = "";

            for (int i = 0; i < dataGridViewProduct.Rows.Count; i++)
            {
                sTemp = dataGridViewProduct.Rows[i].Cells[1].Value.ToString();
                if (sTemp.IndexOf(textBoxMC.Text.Trim()) != -1)
                {
                    iRow = i;
                    break;
                }
            }


            if (iRow != -1)
            {
                //dataGridViewDWLB.Rows[iRow].Selected = false;
                dataGridViewProduct.Rows[iRow].Selected = true;
                dataGridViewProduct.FirstDisplayedScrollingRowIndex = iRow;
            }
            else
            {
                if (dataGridViewProduct.Rows.Count > 0)
                {
                    dataGridViewProduct.Rows[0].Selected = true;
                    dataGridViewProduct.FirstDisplayedScrollingRowIndex = 0;
                }
            }

        }


        protected override bool ProcessCmdKey(ref   Message msg, Keys keyData)
        {
            if (keyData == Keys.F9)
            {
                btnAll_Click(null, null);
                return true;
            }
            if (keyData == Keys.F7)
            {
                btnSearch_Click(null, null);
                return true;
            }
            if (keyData == Keys.F8)
            {
                btnLocation_Click(null, null);
                return true;
            }

            return base.ProcessCmdKey(ref msg, keyData);
        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            int i, j;
            DataSet dsCSV = new DataSet();

            OpenFileDialog openFileDialogOutput = new OpenFileDialog();
            openFileDialogOutput.Filter = "EXCEL files(*.xlsx)|*.xlsx|EXCEL files(*.xls)|*.xls";//
            openFileDialogOutput.FilterIndex = 0;
            openFileDialogOutput.RestoreDirectory = true;

            if (openFileDialogOutput.ShowDialog() != DialogResult.OK) return;

            Microsoft.Office.Interop.Excel.Application m_xlsApp = null;
            Workbook m_Workbook = null;
            Worksheet m_Worksheet = null;

            ArrayList al = new ArrayList();

            try
            {
                object objOpt = System.Reflection.Missing.Value;
                m_xlsApp = new Microsoft.Office.Interop.Excel.Application();
                m_Workbook = m_xlsApp.Workbooks.Open(openFileDialogOutput.FileName, objOpt, objOpt, objOpt, objOpt, objOpt, objOpt, objOpt, objOpt, objOpt, objOpt, objOpt, objOpt, objOpt, objOpt);

                m_Worksheet = (Microsoft.Office.Interop.Excel.Worksheet)m_Workbook.Worksheets.get_Item(1); //第一个

                

                if (m_Worksheet.UsedRange.Rows.Count >= 2 && m_Worksheet.UsedRange.Columns.Count >= 4)
                {

                    for (i = 2; i <= m_Worksheet.UsedRange.Rows.Count; i++)
                    {
                        //第二列为CODE
                        if (((Microsoft.Office.Interop.Excel.Range)(m_Worksheet.Cells[i, 2])).ToString() == "")
                        {
                            continue;
                        }
                        cProduct cP = new cProduct();

                        int.TryParse(((Range)(m_Worksheet.Cells[i, 3])).Value2.ToString(),out cP.iNum);
                        if (cP.iNum == 0) cP.iNum = 1;

                        decimal.TryParse(((Range)(m_Worksheet.Cells[i, 4])).Value2.ToString(), out cP.dFR);
                        if (cP.dFR < 0 || cP.dFR >= 100) cP.dFR = 0;

                        cP.sPCode=((Range)(m_Worksheet.Cells[i, 2])).Value2.ToString();

                        if (((Range)(m_Worksheet.Cells[i, 1])) != null && ((Range)(m_Worksheet.Cells[i, 1])).Text.ToString() != "")
                            cP.sPName = ((Range)(m_Worksheet.Cells[i, 1])).Value2.ToString();

                        al.Add(cP);
                    }
                }


            }
            catch (Exception exc)
            {
                MessageBox.Show("Input Fail","Error");
            }
            finally
            {
                m_Worksheet = null;
                m_Workbook = null;
                m_xlsApp.Quit();
                int generation = System.GC.GetGeneration(m_xlsApp);
                m_xlsApp = null;
                System.GC.Collect(generation);
            }

            
            //导入数据库
            sqlConn.Open();

            System.Data.SqlClient.SqlTransaction sqlta = sqlConn.BeginTransaction();
            sqlComm.Transaction = sqlta;
            try
            {
                foreach (cProduct cp in al)
                {
                    sqlComm.CommandText = "SELECT [Product Name], [Product Code] FROM product WHERE UPPER([Product Code])= N'"+cp.sPCode.ToUpper()+"'";
                    sqldr = sqlComm.ExecuteReader();

                    if (sqldr.HasRows)
                    {
                        sqlComm.CommandText = "UPDATE product SET [Product Name] = N'" + cp.sPName + "', [Number of IMEI] = " + cp.iNum.ToString() + ", [Failure Rate] = " + cp.dFR.ToString() + " WHERE UPPER([Product Code])= N'" + cp.sPCode.ToUpper() + "'";
                    }
                    else
                    {
                        sqlComm.CommandText = "INSERT INTO product ([Product Name], [Number of IMEI], [Failure Rate], [Product Code]) VALUES   (N'"+cp.sPName+"', "+cp.iNum.ToString()+", "+cp.dFR.ToString()+", N'"+cp.sPCode+"')";
                    }
                    sqldr.Close();
                    sqlComm.ExecuteNonQuery();

                }
                sqlta.Commit();
                MessageBox.Show("Input Finish", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
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
                initDataView(0);
            }

            

    
        }



    }

    class cProduct
    {

        public string sPCode = "";
        public string sPName = "";
        public int iNum = 1;
        public decimal dFR=0;


        public cProduct() //构造函数
        {
            this.sPCode = "";
            this.sPName = "";
            this.iNum = 1;
            this.dFR = 0;
        }
    }
}