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
    public partial class FormTAC : Form
    {
        private System.Data.SqlClient.SqlConnection sqlConn = new System.Data.SqlClient.SqlConnection();
        private System.Data.SqlClient.SqlCommand sqlComm = new System.Data.SqlClient.SqlCommand();
        private System.Data.SqlClient.SqlDataReader sqldr;
        private System.Data.SqlClient.SqlDataAdapter sqlDA = new System.Data.SqlClient.SqlDataAdapter();
        private System.Data.DataSet dSet = new DataSet();

        public string strConn = "";

        public int intUserLimit = 0;
        public int iRange = 1000000;

        private DataTable dtBuyer = new DataTable();

        public FormTAC()
        {
            InitializeComponent();
        }

        private void toolStripButtonDEL_Click(object sender, EventArgs e)
        {

        }

        private void FormTAC_Load(object sender, EventArgs e)
        {
            sqlConn.ConnectionString = strConn;
            sqlComm.Connection = sqlConn;
            sqlDA.SelectCommand = sqlComm;

            this.Top = 1;
            this.Left = 1;

            initDatatable();
        }

        private void initDatatable()
        {
            int i;

            sqlConn.Open();
            sqlComm.CommandText = "SELECT ID, [TAC Code], [Product ID], [Product Code], [Indentor ID], [Indentor Code], [Init Number] FROM TAC";

            if (dSet.Tables.Contains("TAC")) dSet.Tables["TAC"].Clear();
            sqlDA.Fill(dSet, "TAC");
            sqlConn.Close();

            dataGridViewP.DataSource = dSet.Tables["TAC"];
            for (i = 0; i < dataGridViewP.ColumnCount; i++)
            {
                dataGridViewP.Columns[i].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            }

            dataGridViewP.Columns[0].Visible = false;
            dataGridViewP.Columns[2].Visible = false;
            dataGridViewP.Columns[4].Visible = false;

            setSTAUS();

        }

        private void setSTAUS()
        {
            toolStripStatusLabelC.Text = "Number of TAC:" + dataGridViewP.RowCount.ToString();
        }

        private void ToolStripButtonADD_Click(object sender, EventArgs e)
        {
            FormTAC_CARD formTAC_CARD = new FormTAC_CARD();
            formTAC_CARD.strConn = strConn;
            formTAC_CARD.iStyle = 0;
            formTAC_CARD.iRange = iRange;

            formTAC_CARD.ShowDialog();
            initDatatable();
        }

        private void toolStripButtonEDIT_Click(object sender, EventArgs e)
        {
            FormTAC_CARD formTAC_CARD = new FormTAC_CARD();
            formTAC_CARD.strConn = strConn;
            formTAC_CARD.iStyle = 1;
            formTAC_CARD.iRange = iRange;
            formTAC_CARD.iID = int.Parse(dataGridViewP.SelectedRows[0].Cells[0].Value.ToString());
            formTAC_CARD.iProduct = int.Parse(dataGridViewP.SelectedRows[0].Cells[2].Value.ToString());
            formTAC_CARD.sPCode = dataGridViewP.SelectedRows[0].Cells[3].Value.ToString();
            formTAC_CARD.iIndentor = int.Parse(dataGridViewP.SelectedRows[0].Cells[4].Value.ToString());
            formTAC_CARD.sICode = dataGridViewP.SelectedRows[0].Cells[5].Value.ToString();

            formTAC_CARD.TextBoxTAC.Text = dataGridViewP.SelectedRows[0].Cells[1].Value.ToString();
            formTAC_CARD.numericUpDownNum.Value = int.Parse(dataGridViewP.SelectedRows[0].Cells[6].Value.ToString());

            formTAC_CARD.ShowDialog();
            initDatatable();
        }

        private void printPreviewToolStripButton_Click(object sender, EventArgs e)
        {
            
            string strT = "TAC";
            PrintDGV.Print_DataGridView(dataGridViewP, strT, true, intUserLimit);
        }

        private void printToolStripButton_Click(object sender, EventArgs e)
        {
            string strT = "TAC";
            PrintDGV.Print_DataGridView(dataGridViewP, strT, false, intUserLimit);
        }

        private void toolStripButtonDEL_Click_1(object sender, EventArgs e)
        {
            if (dataGridViewP.SelectedRows.Count < 1)
            {
                MessageBox.Show("please select TAC", "Infomation", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }


            if (MessageBox.Show("do you want to delete TAC？", "Information", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) == DialogResult.No)
                return;

            int i;
            System.Data.SqlClient.SqlTransaction sqlta;

            sqlConn.Open();
            sqlta = sqlConn.BeginTransaction();
            sqlComm.Transaction = sqlta;
            try
            {

                for (i = 0; i < dataGridViewP.SelectedRows.Count; i++)
                {
                    if (dataGridViewP.Rows[i].IsNewRow)
                        continue;

                    sqlComm.CommandText = "DELETE FROM TAC WHERE   (ID = " + dataGridViewP.SelectedRows[i].Cells[0].Value.ToString() + ")";
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
            initDatatable();
        }

        private void btnPCF_Click(object sender, EventArgs e)
        {
            if (textBoxPC.Text.Trim() == "")
                return;
            var q1 = from dt1 in dSet.Tables["TAC"].AsEnumerable()//查询
                     where (dt1.Field<string>(3).Contains(textBoxPC.Text.Trim()))//条件
                     select dt1;
            if (q1.Count() <= 0)
                return;
            DataTable dtBuyer1 = q1.CopyToDataTable<DataRow>();
            dataGridViewP.DataSource = dtBuyer1;
        }

        private void btnICF_Click(object sender, EventArgs e)
        {
            if (textBoxIC.Text.Trim() == "")
                return;
            var q1 = from dt1 in dSet.Tables["TAC"].AsEnumerable()//查询
                     where (dt1.Field<string>(5).Contains(textBoxIC.Text.Trim()))//条件
                     select dt1;
            if (q1.Count() <= 0)
                return;
            DataTable dtBuyer1 = q1.CopyToDataTable<DataRow>();
            dataGridViewP.DataSource = dtBuyer1;
        }

        private void btnAll_Click(object sender, EventArgs e)
        {
            initDatatable();
        }
    }
}
