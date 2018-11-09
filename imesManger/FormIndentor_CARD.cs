﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace imesManger
{
    public partial class FormIndentor_CARD : Form
    {
        private System.Data.SqlClient.SqlConnection sqlConn = new System.Data.SqlClient.SqlConnection();
        private System.Data.SqlClient.SqlCommand sqlComm = new System.Data.SqlClient.SqlCommand();
        private System.Data.SqlClient.SqlDataReader sqldr;
        private System.Data.SqlClient.SqlDataAdapter sqlDA = new System.Data.SqlClient.SqlDataAdapter();
        private System.Data.DataSet dSet = new DataSet();

        public string strConn = "";

        public int iStyle = 0;
        public DataTable dt;
        public int iSelect = 0;

        public FormIndentor_CARD()
        {
            InitializeComponent();
        }

        private void FormIndentor_CARD_Load(object sender, EventArgs e)
        {

            if (dt.Rows.Count < 1 && iStyle==1)
            {
                this.Close();
                return;
            }


            sqlConn.ConnectionString = strConn;
            sqlComm.Connection = sqlConn;
            sqlDA.SelectCommand = sqlComm;

            switch (iStyle)
            {
                case 0://增加
                    btnAccept.Text = "Add";
                    break;
                case 1://修改
                    btnAccept.Text = "Edit";
                    break;
                default:
                    break;
            }



            if (iStyle == 1) //修改
            {
                //name code
                textBoxDWBH.Text = dt.Rows[0][2].ToString();
                textBoxDWMC.Text = dt.Rows[0][1].ToString();
            }
            



        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            iSelect = 0;
            this.Close();
        }


        private bool countAmount()
        {
            bool bCheck = true;

            if (textBoxDWBH.ToString() == "")
            {
                MessageBox.Show("please input code", "infomation", MessageBoxButtons.OK, MessageBoxIcon.Information);
                bCheck = false;
                return bCheck;
            }

            //if (textBoxDWMC.ToString() == "")
            //{
            //    MessageBox.Show("please input name", "infomation", MessageBoxButtons.OK, MessageBoxIcon.Information);
            //    bCheck = false;
            //}
            return bCheck;
        }

        private void btnAccept_Click(object sender, EventArgs e)
        {
            System.Data.SqlClient.SqlTransaction sqlta;

            if (!countAmount())
            {
               return;
            }
            

            switch (iStyle)
            {
                case 0://增加
                    sqlConn.Open();

                    //查重
                    if (textBoxDWBH.Text.Trim() == "")
                    {
                        MessageBox.Show("please input code");
                        sqlConn.Close();
                        break;
                    }
                    sqlComm.CommandText = "SELECT [Indentor Name], [Indentor Code] FROM indentor WHERE [Indentor Code]= '" + textBoxDWBH.Text.Trim() + "'";
                    sqldr = sqlComm.ExecuteReader();

                    if (sqldr.HasRows)
                    {
                        sqldr.Read();
                        MessageBox.Show("indentor code" + textBoxDWBH.Text.Trim() + "duplicate，name is：" + sqldr.GetValue(1).ToString());
                        sqldr.Close();
                        sqlConn.Close();
                        break;
                    }
                    sqldr.Close();

                    sqlta = sqlConn.BeginTransaction();
                    sqlComm.Transaction = sqlta;
                    try
                    {

                        //得到表单号

                        sqlComm.CommandText = "INSERT INTO indentor ([Indentor Name], [Indentor Code]) VALUES (N'" + textBoxDWMC.Text.Trim() + "', N'" + textBoxDWBH.Text.Trim() + "')"; sqlComm.ExecuteNonQuery();

                        sqlComm.CommandText = "SELECT @@IDENTITY";
                        sqldr = sqlComm.ExecuteReader();
                        sqldr.Read();
                        iSelect = Convert.ToInt32(sqldr.GetValue(0).ToString());
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
                    MessageBox.Show("add finished", "infomation", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    this.Close();
                    break;
                case 1://修改

                    sqlConn.Open();
                    //查重
                    if (textBoxDWBH.Text.Trim() == "")
                    {
                        MessageBox.Show("please input code");
                        sqlConn.Close();
                        break;
                    }
                    iSelect = Convert.ToInt32(dt.Rows[0][0].ToString());
                    sqlComm.CommandText = "SELECT ID, [Indentor Name] FROM indentor WHERE ([Indentor Code] = '" + textBoxDWBH.Text.Trim() + "' AND ID <> " + iSelect.ToString() + ")";
                    sqldr = sqlComm.ExecuteReader();

                    if (sqldr.HasRows)
                    {
                        sqldr.Read();
                        MessageBox.Show("Indentor code " + textBoxDWBH.Text.Trim() + "duplicate，name is：" + sqldr.GetValue(1).ToString());
                        sqldr.Close();
                        sqlConn.Close();
                        break;
                    }
                    sqldr.Close();



                    sqlta = sqlConn.BeginTransaction();
                    sqlComm.Transaction = sqlta;
                    try
                    {

                        iSelect = Convert.ToInt32(dt.Rows[0][0].ToString());
                        sqlComm.CommandText = "UPDATE indentor SET [Indentor Name] = N'" + textBoxDWMC.Text.Trim() + "', [Indentor Code] = N'" + textBoxDWBH.Text.Trim() + "' WHERE (ID = " + iSelect + ")";
                        sqlComm.ExecuteNonQuery();

                        sqlComm.CommandText = "UPDATE acquire SET [Indentor Code] = N'" + textBoxDWBH.Text.Trim() + "' WHERE ([Indentor ID] = " + iSelect + ")";
                        sqlComm.ExecuteNonQuery();

                        sqlComm.CommandText = "UPDATE actual SET [Indentor Code] = N'" + textBoxDWBH.Text.Trim() + "' WHERE ([Indentor ID] = " + iSelect + ")";
                        sqlComm.ExecuteNonQuery();

                        sqlComm.CommandText = "UPDATE TAC SET [Indentor Code] = N'" + textBoxDWBH.Text.Trim() + "' WHERE ([Indentor ID] = " + iSelect + ")";
                        sqlComm.ExecuteNonQuery();

                        sqlComm.CommandText = "UPDATE orders SET [Indentor Code] = N'" + textBoxDWBH.Text.Trim() + "' WHERE ([Indentor ID] = " + iSelect + ")";
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
                    MessageBox.Show("edit finished", "infomation", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    this.Close();
                    break;
                default:
                    break;
            }
        }






   }
}