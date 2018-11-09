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
    public partial class FormTAC_CARD : Form
    {
        private System.Data.SqlClient.SqlConnection sqlConn = new System.Data.SqlClient.SqlConnection();
        private System.Data.SqlClient.SqlCommand sqlComm = new System.Data.SqlClient.SqlCommand();
        private System.Data.SqlClient.SqlDataReader sqldr;
        private System.Data.SqlClient.SqlDataAdapter sqlDA = new System.Data.SqlClient.SqlDataAdapter();
        private System.Data.DataSet dSet = new DataSet();

        public string strConn = "";

        public int iStyle = 0; //0增加 1修改
        public DataTable dt;
        public int iSelect = 0;

        public int iProduct = 0;
        public string sPCode = "";
        public int iIndentor = 0;
        public string sICode = "";

        public int iID = 0;

        public int iRange = 1000000;

        public FormTAC_CARD()
        {
            InitializeComponent();
        }

        private void FormTAC_CARD_Load(object sender, EventArgs e)
        {
            sqlConn.ConnectionString = strConn;
            sqlComm.Connection = sqlConn;
            sqlDA.SelectCommand = sqlComm;

            switch (iStyle)
            {
                case 0://增加
                    btnAccept.Text = "Add";
                    numericUpDownNum.ReadOnly = false;
                    break;
                case 1://修改
                    btnAccept.Text = "Edit";
                    numericUpDownNum.ReadOnly = true;
                    break;
                default:
                    break;
            }

            initDatatable();
            numericUpDownNum.Increment = iRange;

        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void initDatatable()
        {
            int i;

            sqlConn.Open();
            sqlComm.CommandText = "SELECT ID, [Product Name], [Product Code] FROM product";
            if (iStyle == 1)
            {
                sqlComm.CommandText += " WHERE (ID = "+iProduct+")";
            }

            if (dSet.Tables.Contains("product")) dSet.Tables["product"].Clear();
            sqlDA.Fill(dSet, "product");

            sqlComm.CommandText = "SELECT   ID, [Indentor Name], [Indentor Code] FROM indentor";
            if (iStyle == 1)
            {
                sqlComm.CommandText += " WHERE (ID = " + iIndentor + ")";
            }

            if (dSet.Tables.Contains("indentor")) dSet.Tables["indentor"].Clear();
            sqlDA.Fill(dSet, "indentor");


            sqlConn.Close();

            dataGridViewP.SelectionChanged -= dataGridViewP_SelectionChanged;
            dataGridViewI.SelectionChanged -= dataGridViewI_SelectionChanged;

            dataGridViewP.DataSource = dSet.Tables["product"];
            dataGridViewP.Columns[2].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            dataGridViewP.Columns[0].Visible = false;
            dataGridViewP.Columns[1].Visible = false;

            dataGridViewI.DataSource = dSet.Tables["indentor"];
            dataGridViewI.Columns[2].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            dataGridViewI.Columns[0].Visible = false;
            dataGridViewI.Columns[1].Visible = false;

            dataGridViewP_SelectionChanged(null,null);
            dataGridViewI_SelectionChanged(null, null);
            dataGridViewP.SelectionChanged += dataGridViewP_SelectionChanged;
            dataGridViewI.SelectionChanged += dataGridViewI_SelectionChanged;

        }

        private void dataGridViewP_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridViewP.Rows.Count < 1)
                return;

            if (iStyle == 0) //新增
            {
                labelP.Text = dataGridViewP.SelectedRows[0].Cells[1].Value.ToString();
                iProduct = int.Parse(dataGridViewP.SelectedRows[0].Cells[0].Value.ToString());
                sPCode = dataGridViewP.SelectedRows[0].Cells[2].Value.ToString();
            }
            else //修改
            {
                labelP.Text = dataGridViewP.SelectedRows[0].Cells[1].Value.ToString();
            }

        }

        private void dataGridViewI_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridViewI.Rows.Count < 1)
                return;

            if (iStyle == 0) //新增
            {
                labelI.Text = dataGridViewI.SelectedRows[0].Cells[1].Value.ToString();
                iIndentor = int.Parse(dataGridViewI.SelectedRows[0].Cells[0].Value.ToString());
                sICode = dataGridViewI.SelectedRows[0].Cells[2].Value.ToString();

                getLastNumber();
            }
            else //修改
            {
                labelI.Text = dataGridViewI.SelectedRows[0].Cells[1].Value.ToString();
            }

            
        }

        private void getLastNumber() //得到开始区间数
        {

            sqlComm.CommandText="SELECT MAX([Init Number]) AS MAXNUMBER FROM TAC WHERE ([Product ID] = "+iProduct.ToString()+") AND ([Indentor ID] = "+iIndentor+")";

            sqlConn.Open();
            sqldr = sqlComm.ExecuteReader();
            if (sqldr.HasRows)
            {
                sqldr.Read();
                if (sqldr.GetValue(0).ToString() == "")
                {
                    numericUpDownNum.Value = 0;
                }
                else
                {
                    numericUpDownNum.Value = int.Parse(sqldr.GetValue(0).ToString())+iRange;
                }

            }
            else
            {
                numericUpDownNum.Value = 0;
            }
            sqldr.Close();
            sqlConn.Close();
        }

        private void btnAccept_Click(object sender, EventArgs e)
        {
            System.Data.SqlClient.SqlTransaction sqlta;
            if(TextBoxTAC.Text.Trim()=="")
            {
                MessageBox.Show("please input TAC code");
                return;
            }

            switch (iStyle)
            {
                case 0://增加
                    sqlConn.Open();

                    sqlComm.CommandText = "SELECT [TAC Code], [Product Code], [Indentor Code] FROM TAC WHERE ([TAC Code] = N'" + TextBoxTAC.Text.Trim() + "')";
                    sqldr = sqlComm.ExecuteReader();

                    if (sqldr.HasRows)
                    {
                        sqldr.Read();
                        MessageBox.Show("TAC code" + TextBoxTAC.Text.Trim() + "duplicate，code is：" + sqldr.GetValue(1).ToString() + " ," + sqldr.GetValue(2).ToString());
                        sqldr.Close();
                        sqlConn.Close();
                        break;
                    }
                    sqldr.Close();

                    sqlta = sqlConn.BeginTransaction();
                    sqlComm.Transaction = sqlta;
                    try
                    {
                        //得到ID号
                        sqlComm.CommandText = "INSERT INTO TAC ([TAC Code], [Product ID], [Product Code], [Indentor ID], [Indentor Code], [Init Number]) VALUES   (N'" + TextBoxTAC.Text.Trim() + "', "+iProduct.ToString()+", N'"+sPCode+"', "+iIndentor.ToString()+", N'"+sICode+"', "+numericUpDownNum.Value.ToString()+")";
                        sqlComm.ExecuteNonQuery();

                        sqlComm.CommandText = "SELECT @@IDENTITY";
                        sqldr = sqlComm.ExecuteReader();
                        sqldr.Read();
                        iID = Convert.ToInt32(sqldr.GetValue(0).ToString());
                        sqldr.Close();


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
                    MessageBox.Show("add TAC CODE finished", "infomation", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    this.Close();
                    break;
                case 1://修改
                    sqlConn.Open();
                    //查重
                    sqlComm.CommandText = "SELECT [TAC Code], [Product Code], [Indentor Code] FROM TAC WHERE ([TAC Code] = N'" + TextBoxTAC.Text.Trim() + "' AND ID <> " + iID.ToString() + ")";
                    sqldr = sqlComm.ExecuteReader();

                    if (sqldr.HasRows)
                    {
                        sqldr.Read();
                        MessageBox.Show("TAC code" + TextBoxTAC.Text.Trim() + "duplicate，code is：" + sqldr.GetValue(1).ToString() + " ," + sqldr.GetValue(2).ToString());
                        sqldr.Close();
                        sqlConn.Close();
                        break;
                    }
                    sqldr.Close();

                    sqlta = sqlConn.BeginTransaction();
                    sqlComm.Transaction = sqlta;
                    try
                    {

                        sqlComm.CommandText = "UPDATE  TAC SET [TAC Code] = N'" + TextBoxTAC.Text.Trim() + "' WHERE (ID = " + iID.ToString() + ")";
                        sqlComm.ExecuteNonQuery();

                        sqlComm.CommandText = "UPDATE acquire SET [TAC Code] = N'" + TextBoxTAC.Text.Trim() + "' WHERE ([TAC ID] = " + iID.ToString() + ")";
                        sqlComm.ExecuteNonQuery();

                        sqlComm.CommandText = "UPDATE actual SET [TAC Code] = N'" + TextBoxTAC.Text.Trim() + "' WHERE ([TAC ID] = " + iID.ToString() + ")";
                        sqlComm.ExecuteNonQuery();

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
                    MessageBox.Show("edit TAC CODE finished", "infomation", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    this.Close();
                    break;

            }
        }
    }
}
