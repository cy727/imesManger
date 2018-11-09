using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Xml;
using System.IO;


namespace imesManger
{
    public partial class formDatabaseSet : Form
    {
        public string strConn="";
        private string dFileName="";
        public int intMode = 0;

        private System.Data.DataSet dSet = new DataSet();

        public formDatabaseSet()
        {
            InitializeComponent();
            dFileName = Directory.GetCurrentDirectory() + "\\appcon.xml";
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            strConn = "";
            this.Close();
        }

        private void btnTest_Click(object sender, EventArgs e)
        {
            
            strConn = "workstation id=CY;packet size=4096;user id=" + textBoxUser.Text.Trim() + ";password=" + textBoxPassword.Text.Trim() + ";data source=\"" + textBoxServer.Text.Trim() + "\";;initial catalog="+textBoxDatabase.Text.Trim();

            sqlConn.ConnectionString = strConn;
            try
            {
                sqlConn.Open();
            }
            catch (System.Data.SqlClient.SqlException err)
            {
                MessageBox.Show("connect fail");
                strConn = "";
                return;

            }

            MessageBox.Show("connect sucess");
            sqlConn.Close();

            dSet.Tables["DataBaseInfo"].Rows[0][0] = textBoxServer.Text;
            dSet.Tables["DataBaseInfo"].Rows[0][1] = textBoxUser.Text;

            if(checkBoxRember.Checked) //记住密码
                dSet.Tables["DataBaseInfo"].Rows[0][2] = textBoxPassword.Text;
            else
                dSet.Tables["DataBaseInfo"].Rows[0][2] = "";

            dSet.Tables["DataBaseInfo"].Rows[0][3] = textBoxDatabase.Text;
            dSet.WriteXml(dFileName);


            this.Close();

        }

        private void formDatabaseSet_Load(object sender, EventArgs e)
        {
            if (intMode == 0)//测试
            {
                btnTest.Visible = true;
                btnCreate.Visible = false;
                this.Text = "Set";
            }
            else //创建
            {
                btnTest.Visible = false;
                btnCreate.Visible = true;
                this.Text = "Create";
            }

            sqlComm.Connection = sqlConn;

            if(File.Exists(dFileName)) //存在文件
            {
                dSet.ReadXml(dFileName);
            }
            else  //建立文件
            {
                dSet.Tables.Add("DataBaseInfo");

                dSet.Tables["DataBaseInfo"].Columns.Add("Address", System.Type.GetType("System.String"));
                dSet.Tables["DataBaseInfo"].Columns.Add("User", System.Type.GetType("System.String"));
                dSet.Tables["DataBaseInfo"].Columns.Add("Pass", System.Type.GetType("System.String"));
                dSet.Tables["DataBaseInfo"].Columns.Add("DB", System.Type.GetType("System.String"));

                string[]  strDRow ={ "","","",""};
                dSet.Tables["DataBaseInfo"].Rows.Add(strDRow);
            }

            textBoxServer.Text = dSet.Tables["DataBaseInfo"].Rows[0][0].ToString();
            textBoxUser.Text = dSet.Tables["DataBaseInfo"].Rows[0][1].ToString();
            textBoxPassword.Text = dSet.Tables["DataBaseInfo"].Rows[0][2].ToString();
            textBoxDatabase.Text = dSet.Tables["DataBaseInfo"].Rows[0][3].ToString();
        }

        private void btnCreate_Click(object sender, EventArgs e)
        {


            strConn = "packet size=4096;user id=" + textBoxUser.Text.Trim() + ";password=" + textBoxPassword.Text.Trim() + ";data source=\"" + textBoxServer.Text.Trim() + "\";Database=" + textBoxDatabase.Text.Trim();
            try
            {
                sqlConn.ConnectionString = strConn;
                sqlConn.Open();
                sqlConn.Close();

                strConn = "packet size=4096;user id=" + textBoxUser.Text.Trim() + ";password=" + textBoxPassword.Text.Trim() + ";data source=\"" + textBoxServer.Text.Trim() + "\";Database=master";

                if (MessageBox.Show("exist database " + textBoxDatabase.Text.Trim() + " ,delete it?", "Information", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) == DialogResult.Yes)
                {
                    sqlConn.ConnectionString = strConn;
                    sqlConn.Open();


                    sqlComm.CommandText = @"ALTER DATABASE " + textBoxDatabase.Text.Trim() + @" SET SINGLE_USER WITH ROLLBACK IMMEDIATE; DROP DATABASE [" + textBoxDatabase.Text.Trim() + "]";
                    //sqlComm.ExecuteNonQuery();
                    //sqlComm.CommandText = "drop database " + textBoxDatabase.Text.Trim();
                    sqlComm.ExecuteNonQuery();
                }


            }
            catch (Exception ex)
            {
                //int iii = 0;
            }
            finally
            {
                sqlConn.Close();
            }


            strConn = "packet size=4096;user id=" + textBoxUser.Text.Trim() + ";password=" + textBoxPassword.Text.Trim() + ";data source=\"" + textBoxServer.Text.Trim() + "\";initial catalog=master;Integrated Security=True";

            strConn = "workstation id=CY;packet size=4096;user id=" + textBoxUser.Text.Trim() + ";password=" + textBoxPassword.Text.Trim() + ";data source=\"" + textBoxServer.Text.Trim() + "\";;initial catalog=master" ;


            try
            {
                sqlConn.ConnectionString = strConn;
                sqlConn.Open();
            }
            catch (Exception ex)
            {
                MessageBox.Show("create fail：" + ex.Message.ToString(), "error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                strConn = "";
                return;
            }

            try

            {
                sqlComm.CommandText = "create database " + textBoxDatabase.Text.Trim();
                sqlComm.ExecuteNonQuery();

                sqlConn.Close();

                //strConn = "packet size=4096;user id=" + textBoxUser.Text.Trim() + ";password=" + textBoxPassword.Text.Trim() + ";data source=\"" + textBoxServer.Text.Trim() + "\";Database=" + textBoxDatabase.Text.Trim() + ";Integrated Security=SSPI";
                strConn = "packet size=4096;user id=" + textBoxUser.Text.Trim() + ";password=" + textBoxPassword.Text.Trim() + ";data source=\"" + textBoxServer.Text.Trim() + "\";Database=" + textBoxDatabase.Text.Trim();
                sqlConn.ConnectionString = strConn;
                sqlConn.Open();

                sqlComm.CommandText = "CREATE TABLE [dbo].[acquire]([ID] [int] IDENTITY(1,1) NOT NULL, [Product ID] [int] NULL,[Product Code] [nvarchar](50) NULL,[Indentor ID] [int] NULL,[Indentor Code] [nvarchar](50) NULL,[Start Number] [int] NULL,[End Number] [int] NULL,[Date] [smalldatetime] NULL,[Year] [int] NULL,[Num of Week] [nvarchar](50) NULL,[TAC ID] [int] NULL,[TAC Code] [nvarchar](50) NULL,[Status] [smallint] NULL,CONSTRAINT [PK_order] PRIMARY KEY CLUSTERED ([ID] ASC) WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]) ON [PRIMARY]";
                sqlComm.ExecuteNonQuery();

                sqlComm.CommandText = "CREATE TABLE [dbo].[actual]([ID] [int] IDENTITY(1,1) NOT NULL,[Product ID] [int] NULL,[Product Code] [nvarchar](50) NULL,[Indentor ID] [int] NULL,[Indentor Code] [nvarchar](50) NULL,[Start Number] [int] NULL,[End Number] [int] NULL,[Acquire ID] [int] NULL,[Date] [smalldatetime] NULL,[Year] [int] NULL,[Num of Week] [nvarchar](50) NULL,[TAC ID] [int] NULL,[TAC Code] [nvarchar](50) NULL,[Total Number] [int] NULL,CONSTRAINT [PK_actual] PRIMARY KEY CLUSTERED ([ID] ASC)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]) ON [PRIMARY]";
                sqlComm.ExecuteNonQuery();

                sqlComm.CommandText = "CREATE TABLE [dbo].[indentor]([ID] [int] IDENTITY(1,1) NOT NULL,[Indentor Name] [nvarchar](150) NULL,[Indentor Code] [nvarchar](50) NULL,CONSTRAINT [PK_indentor] PRIMARY KEY CLUSTERED ([ID] ASC)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]) ON [PRIMARY]";
                sqlComm.ExecuteNonQuery();

                sqlComm.CommandText = "CREATE TABLE [dbo].[orders]([ID] [int] IDENTITY(1,1) NOT NULL,[Product ID] [int] NULL,[Product Code] [nvarchar](50) NULL,[Indentor ID] [int] NULL,[Indentor Code] [nvarchar](50) NULL,[Start Number] [int] NULL,[End Number] [int] NULL,[Acquire ID] [int] NULL,[Date] [smalldatetime] NULL,[Year] [int] NULL,[Num of Week] [nvarchar](50) NULL,[Numbers] [int] NULL,[Order Start Number] [int] NULL,[Order End Number] [int] NULL,[backlog] [int] NULL,CONSTRAINT [PK_orders] PRIMARY KEY CLUSTERED ([ID] ASC)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]) ON [PRIMARY]";
                sqlComm.ExecuteNonQuery();

                sqlComm.CommandText = "CREATE TABLE [dbo].[parameters]([ID] [int] IDENTITY(1,1) NOT NULL,[num of week] [int] NULL,[num of pass] [int] NULL,[range] [int] NULL,[string TAG] [nvarchar](50) NULL,[SAP Control Number] [int] NULL,[SAP order] [nvarchar](50) NULL,[Sales plan] [nvarchar](50) NULL,[CW IMEI] [nvarchar](50) NULL,[Master data] [nvarchar](50) NULL,CONSTRAINT [PK_parameters] PRIMARY KEY CLUSTERED ([ID] ASC)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]) ON [PRIMARY]";
                sqlComm.ExecuteNonQuery();

                sqlComm.CommandText = "CREATE TABLE [dbo].[product]([ID] [int] IDENTITY(1,1) NOT NULL,[Product Name] [nvarchar](100) NULL,[Product Code] [nvarchar](255) NULL,[Number of IMEI] [int] NULL,[Status] [bit] NULL,[Factory ID] [int] NULL,[Factory Code] [nvarchar](100) NULL,[Failure Rate] [decimal](10, 0) NULL) ON [PRIMARY]";
                sqlComm.ExecuteNonQuery();

                sqlComm.CommandText = "CREATE TABLE [dbo].[TAC]([ID] [int] IDENTITY(1,1) NOT NULL,[TAC Code] [nvarchar](50) NULL,[Product ID] [int] NULL,[Product Code] [nvarchar](50) NULL,[Indentor ID] [int] NULL,[Indentor Code] [nvarchar](50) NULL,[Init Number] [int] NOT NULL) ON [PRIMARY]";
                sqlComm.ExecuteNonQuery();

                sqlComm.CommandText = "ALTER TABLE [dbo].[TAC] ADD  CONSTRAINT [DF_TAC_Init Number]  DEFAULT ((0)) FOR [Init Number]";
                sqlComm.ExecuteNonQuery();

                sqlComm.CommandText = "INSERT INTO parameters([num of week], [num of pass], range, [string TAG], [SAP Control Number], [SAP order], [Sales plan], [CW IMEI],  [Master data]) VALUES   (11, 2, 1000000, N'Distrubutor', 4, N'SAP order', N'Sales plan', N'CW IMEI', N'Master')";
                sqlComm.ExecuteNonQuery();

                MessageBox.Show("Create Database Success" , "Database", MessageBoxButtons.OK, MessageBoxIcon.Information);

                dSet.Tables["DataBaseInfo"].Rows[0][0] = textBoxServer.Text;
                dSet.Tables["DataBaseInfo"].Rows[0][1] = textBoxUser.Text;

                if (checkBoxRember.Checked) //记住密码
                    dSet.Tables["DataBaseInfo"].Rows[0][2] = textBoxPassword.Text;
                else
                    dSet.Tables["DataBaseInfo"].Rows[0][2] = "";

                dSet.Tables["DataBaseInfo"].Rows[0][3] = textBoxDatabase.Text;
                dSet.WriteXml(dFileName);
                

                
            }
            catch (Exception ex)
            {
                MessageBox.Show("Create fail：" + ex.Message.ToString(), "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                strConn = "";
            }
            finally
            {
                sqlConn.Close();
                this.Dispose();
            }
         

        }
    }
}