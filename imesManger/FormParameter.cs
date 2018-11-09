using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;

namespace imesManger
{
    public partial class FormParameter : Form
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

        string dFileName = Directory.GetCurrentDirectory() + "\\options.xml";

        public FormParameter()
        {
            InitializeComponent();
        }

        private void FormParameter_Load(object sender, EventArgs e)
        {
            object[] oTemp = new object[3];

            if (strConn != "")
            {
                sqlConn.ConnectionString = strConn;
                sqlComm.Connection = sqlConn;
                sqlDA.SelectCommand = sqlComm;
            }
            else
            {
                if (File.Exists(dFileName)) //存在文件
                {
                    dSet.ReadXml(dFileName);

                    numericUpDownWeek.Value = int.Parse(dSet.Tables["parameters"].Rows[0][0].ToString());
                    numericUpDownPass.Value = int.Parse(dSet.Tables["parameters"].Rows[0][1].ToString());
                    numericUpDownTAC.Value = int.Parse(dSet.Tables["parameters"].Rows[0][2].ToString());
                    
                }
                else //没有文件，建立datatable
                {
                    dSet.Tables.Add("parameters");
                    dSet.Tables["parameters"].Columns.Add("Weeks", System.Type.GetType("System.Decimal"));
                    dSet.Tables["parameters"].Columns.Add("Pass", System.Type.GetType("System.Decimal"));
                    dSet.Tables["parameters"].Columns.Add("TAC", System.Type.GetType("System.Decimal"));

                    oTemp[0] = numericUpDownWeek.Value;
                    oTemp[1] = numericUpDownPass.Value;
                    oTemp[2] = numericUpDownTAC.Value;

                    dSet.Tables["parameters"].Rows.Add(oTemp);

                }
            }


        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            object[] oTemp = new object[3];
            if (strConn != "")
            {
                sqlConn.Open();
                sqlComm.CommandText = "UPDATE parameters SET [num of week] =" + numericUpDownWeek.Value.ToString() + ", [num of pass] =" + numericUpDownPass.Value.ToString();
                sqlComm.ExecuteNonQuery();
                sqlConn.Close();
            }
            else
            {
                dSet.Tables["parameters"].Rows.Clear();
                oTemp[0] = numericUpDownWeek.Value;
                oTemp[1] = numericUpDownPass.Value;
                oTemp[2] = numericUpDownTAC.Value;

                dSet.Tables["parameters"].Rows.Add(oTemp);
                dSet.WriteXml(dFileName);
            }
            this.Close();

        }
    }
}
