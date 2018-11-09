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
using System.Linq;

namespace imesManger
{
    public partial class FormBuyer : Form
    {
        private System.Data.SqlClient.SqlConnection sqlConn = new System.Data.SqlClient.SqlConnection();
        private System.Data.SqlClient.SqlCommand sqlComm = new System.Data.SqlClient.SqlCommand();
        private System.Data.SqlClient.SqlDataReader sqldr;
        private System.Data.SqlClient.SqlDataAdapter sqlDA = new System.Data.SqlClient.SqlDataAdapter();
        private System.Data.DataSet dSet = new DataSet();

        public string strConn = "";

        public int intUserLimit = 0;

        private DataTable dtBuyer = new DataTable();


        public FormBuyer()
        {
            InitializeComponent();
        }

        private void FormBuyer_Load(object sender, EventArgs e)
        {
            sqlConn.ConnectionString = strConn;
            sqlComm.Connection = sqlConn;
            sqlDA.SelectCommand = sqlComm;

            this.Top = 1;
            this.Left = 1;

            dtBuyer.Columns.Add("pid", System.Type.GetType("System.Decimal"));//1
            dtBuyer.Columns.Add("Product Name", System.Type.GetType("System.String"));
            dtBuyer.Columns.Add("Product Code", System.Type.GetType("System.String"));
            dtBuyer.Columns.Add("iid", System.Type.GetType("System.Decimal"));//1
            dtBuyer.Columns.Add("Indentor Name", System.Type.GetType("System.String"));
            dtBuyer.Columns.Add("Indentor Code", System.Type.GetType("System.String"));
            dtBuyer.Columns.Add("Pre Code", System.Type.GetType("System.String"));

            initDatatable();
        }
        
        private void initDatatable()
        {
            sqlConn.Open();
            sqlComm.CommandText = "SELECT product.ID, product.[Product Name], product.[Product Code], indentor.ID AS iID, indentor.[Indentor Name], indentor.[Indentor Code] FROM product CROSS JOIN indentor";

            if (dSet.Tables.Contains("pi")) dSet.Tables["pi"].Clear();
            sqlDA.Fill(dSet, "pi");

            sqlComm.CommandText = "SELECT ID, [Product ID], [Product Code], [Indentor ID], [Indentor Code], [Pre Code], [Current ID], [Current Count],  [Order ID], [Order Count] FROM buyer";

            if (dSet.Tables.Contains("buyer")) dSet.Tables["buyer"].Clear();
            sqlDA.Fill(dSet, "buyer");

            sqlConn.Close();

            object[] oTemp = new object[7];
            dtBuyer.Clear();

            int i;
            for (i = 0; i < dSet.Tables["pi"].Rows.Count; i++)
            {
                oTemp[0] = dSet.Tables["pi"].Rows[i][0];
                oTemp[1] = dSet.Tables["pi"].Rows[i][1];
                oTemp[2] = dSet.Tables["pi"].Rows[i][2];
                oTemp[3] = dSet.Tables["pi"].Rows[i][3];
                oTemp[4] = dSet.Tables["pi"].Rows[i][4];
                oTemp[5] = dSet.Tables["pi"].Rows[i][5];
                oTemp[6] = "";

                var q1 = from dt1 in dSet.Tables["buyer"].AsEnumerable()//查询
                         where (dt1.Field<int>("Product ID") == int.Parse(oTemp[0].ToString())) && (dt1.Field<int>("Indentor ID") == int.Parse(oTemp[3].ToString()))//条件
                         select dt1;

                foreach (var item in q1)//显示查询结果
                {
                    oTemp[6] = item.Field<string>("Pre Code");
                    break;
                }

                dtBuyer.Rows.Add(oTemp);
            }

            dataGridViewP.DataSource = dtBuyer;
            for (i = 0; i < dataGridViewP.ColumnCount-1; i++)
            {
                dataGridViewP.Columns[i].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                dataGridViewP.Columns[i].ReadOnly = true;
            }
            dataGridViewP.Columns[i].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;

            dataGridViewP.Columns[0].Visible = false;
            dataGridViewP.Columns[3].Visible = false;

        }

        private void toolStripButtonEDIT_Click(object sender, EventArgs e)
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
                    //if has this buyer
                    var q1 = from dt1 in dSet.Tables["buyer"].AsEnumerable()//查询
                             where (dt1.Field<int>("Product ID") == int.Parse(dataGridViewP.Rows[i].Cells[0].Value.ToString())) && (dt1.Field<int>("Indentor ID") == int.Parse(dataGridViewP.Rows[i].Cells[0].Value.ToString()))//条件
                             select dt1;

                    if (q1.Count() > 0) //has buyer already
                    {
                        sqlComm.CommandText = "UPDATE buyer SET [Pre Code] = N'" + dataGridViewP.Rows[i].Cells[6].Value.ToString() + "' WHERE ([Product ID] = " + dataGridViewP.Rows[i].Cells[0].Value.ToString() + ") AND ([Indentor ID] = " + dataGridViewP.Rows[i].Cells[3].Value.ToString() + ")"; 
                    }
                    else //has not buyer already
                    {
                        sqlComm.CommandText = "INSERT INTO buyer ([Product ID], [Product Code], [Indentor ID], [Indentor Code], [Pre Code], [Current ID], [Current Count], [Order ID],  [Order Count]) VALUES (" + dataGridViewP.Rows[i].Cells[0].Value.ToString() + ", N'" + dataGridViewP.Rows[i].Cells[2].Value.ToString() + "', " + dataGridViewP.Rows[i].Cells[3].Value.ToString() + ", N'" + dataGridViewP.Rows[i].Cells[5].Value.ToString() + "', N'" + dataGridViewP.Rows[i].Cells[6].Value.ToString() + "', 0, 0, N'0', 0)";
                    }
                    sqlComm.ExecuteNonQuery();



                }
                sqlta.Commit();

                MessageBox.Show("Edit Finished", "information",MessageBoxButtons.OK, MessageBoxIcon.Information);
                
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
                initDatatable();
            }
        }

        private void btnPCF_Click(object sender, EventArgs e)
        {
            if (textBoxPC.Text.Trim() == "")
                return;
            var q1 = from dt1 in dtBuyer.AsEnumerable()//查询
                     where (dt1.Field<string>(2) == textBoxPC.Text.Trim())//条件
                     select dt1;
            if (q1.Count() < 0)
                return;
            DataTable dtBuyer1 = q1.CopyToDataTable<DataRow>();
            dataGridViewP.DataSource = dtBuyer1;


        }

        private void btnAll_Click(object sender, EventArgs e)
        {
            initDatatable();
        }

        private void btnICF_Click(object sender, EventArgs e)
        {
            if (textBoxIC.Text.Trim() == "")
                return;
            var q1 = from dt1 in dtBuyer.AsEnumerable()//查询
                     where (dt1.Field<string>(5) == textBoxIC.Text.Trim())//条件
                     select dt1;
            if (q1.Count() < 0)
                return;
            DataTable dtBuyer1 = q1.CopyToDataTable<DataRow>();
            dataGridViewP.DataSource = dtBuyer1;
        }


    }
}
