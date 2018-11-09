using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Xml;
using System.Net;
using System.Data.OleDb;

namespace imesManger
{
    public partial class MDIiManage : Form
    {
        //private int childFormNumber = 0;
        //private System.Data.OleDb.OleDbConnection sqlConn = new System.Data.OleDb.OleDbConnection();
        //private System.Data.OleDb.OleDbCommand sqlComm = new System.Data.OleDb.OleDbCommand();
        //private System.Data.OleDb.OleDbDataReader sqldr;
        //private System.Data.OleDb.OleDbDataAdapter sqlDA = new System.Data.OleDb.OleDbDataAdapter();
        //private System.Data.DataSet dSet = new DataSet();

        private System.Data.SqlClient.SqlConnection sqlConn = new System.Data.SqlClient.SqlConnection();
        private System.Data.SqlClient.SqlCommand sqlComm = new System.Data.SqlClient.SqlCommand();
        private System.Data.SqlClient.SqlDataReader sqldr;
        private System.Data.SqlClient.SqlDataAdapter sqlDA = new System.Data.SqlClient.SqlDataAdapter();
        private System.Data.DataSet dSet = new DataSet();
        private System.Data.DataSet dSetOption = new DataSet();

        private string strDataBaseAddr = "";
        private string strDataBaseUser = "";
        private string strDataBasePass = "";
        private string strDataBaseName = "";

        private string sTAG = "Distrubutor";
        private int iSAPControlNumber = 4;
        private string sSAPorder = "SAP order";
        private string sSalesplan = "Sales plan";
        private string sCWIMEI = "CW IMEI";
        private string sMaster = "Master";

        private string strConn = "";

        private int iNumofWeek = 11;
        private int iNumofPass = 1;
        private int iRange = 1000000;


        public MDIiManage()
        {
            InitializeComponent();
        }

        private void ExitToolsStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void ToolBarToolStripMenuItem_Click(object sender, EventArgs e)
        {
            toolStrip.Visible = toolBarToolStripMenuItem.Checked;
        }

        private void StatusBarToolStripMenuItem_Click(object sender, EventArgs e)
        {
            statusStrip.Visible = statusBarToolStripMenuItem.Checked;
        }

        private void CascadeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            LayoutMdi(MdiLayout.Cascade);
        }

        private void TileVerticalToolStripMenuItem_Click(object sender, EventArgs e)
        {
            LayoutMdi(MdiLayout.TileVertical);
        }

        private void TileHorizontalToolStripMenuItem_Click(object sender, EventArgs e)
        {
            LayoutMdi(MdiLayout.TileHorizontal);
        }

        private void ArrangeIconsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            LayoutMdi(MdiLayout.ArrangeIcons);
        }

        private void CloseAllToolStripMenuItem_Click(object sender, EventArgs e)
        {
            foreach (Form childForm in MdiChildren)
            {
                childForm.Close();
            }
        }

        private void aboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //int iii=ClassIm.GetWeek(new DateTime(2011,1,6));
            new AboutBoxIm().ShowDialog();
        }

        private void exitEToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void MDIiManage_Load(object sender, EventArgs e)
        {
            //日期显示
            toolStripStatusLabelDate.Text = "Today is " + System.DateTime.Now.ToShortDateString() + ", Week " + ClassIm.GetWeek(System.DateTime.Now).ToString();
            
            string dFileName = "";

            dFileName = Directory.GetCurrentDirectory() + "\\appcon.xml";
            sqlComm.Connection = sqlConn;
            sqlDA.SelectCommand = sqlComm;

            if (File.Exists(dFileName)) //存在文件
            {
                dSet.ReadXml(dFileName);
                strConn = "workstation id=CY;packet size=4096;user id=" + dSet.Tables["DataBaseInfo"].Rows[0][1].ToString() + ";password=" + dSet.Tables["DataBaseInfo"].Rows[0][2].ToString() + ";data source=\"" + dSet.Tables["DataBaseInfo"].Rows[0][0].ToString() + "\";;initial catalog=" + dSet.Tables["DataBaseInfo"].Rows[0][3].ToString();

                strDataBaseAddr = dSet.Tables["DataBaseInfo"].Rows[0][0].ToString();
                strDataBaseUser = dSet.Tables["DataBaseInfo"].Rows[0][1].ToString();
                strDataBasePass = dSet.Tables["DataBaseInfo"].Rows[0][2].ToString();
                strDataBaseName = dSet.Tables["DataBaseInfo"].Rows[0][3].ToString();

                if (strDataBaseAddr == "0.0.0.0")
                {
                    strConn = "";

                    //option
                    string dFileNameOption = Directory.GetCurrentDirectory() + "\\options.xml";
                    if (File.Exists(dFileNameOption)) //存在文件
                    {
                        dSetOption.ReadXml(dFileNameOption);

                        iNumofWeek = int.Parse(dSetOption.Tables["parameters"].Rows[0][0].ToString());
                        iNumofPass = int.Parse(dSetOption.Tables["parameters"].Rows[0][1].ToString());
                        //iRange = int.Parse(dSetOption.Tables["parameters"].Rows[0][2].ToString());

                    }
                    return;
                }

                sqlConn.ConnectionString = strConn;
                try
                {
                    sqlConn.Open();
                    sqlDA.SelectCommand = sqlComm;

                    sqlComm.CommandText = "SELECT   ID, [num of week], [num of pass], range, [string TAG], [SAP Control Number], [SAP order], [Sales plan], [CW IMEI], [Master data] FROM parameters";
                    sqldr = sqlComm.ExecuteReader();
                    if (sqldr.HasRows)
                    {
                        sqldr.Read();
                        iNumofWeek = int.Parse(sqldr.GetValue(1).ToString());
                        iNumofPass = int.Parse(sqldr.GetValue(2).ToString());
                        iRange = int.Parse(sqldr.GetValue(3).ToString());

                        sTAG = sqldr.GetValue(4).ToString(); ;
                        iSAPControlNumber = int.Parse(sqldr.GetValue(5).ToString()); ;
                        sSAPorder = sqldr.GetValue(6).ToString();
                        sSalesplan = sqldr.GetValue(7).ToString();
                        sCWIMEI = sqldr.GetValue(8).ToString();
                        sMaster = sqldr.GetValue(9).ToString();
                        sqldr.Close();

                    }
                }
                catch (System.Data.SqlClient.SqlException err)
                {

                    bool isCreateDatabase = true;
                    if (MessageBox.Show("connect database fail, create it？" + err.Message.ToString(), "Infomation", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) == DialogResult.No)
                        isCreateDatabase = false;

                    strConn = "";
                    formDatabaseSet frmDatabaseSet = new formDatabaseSet();
                    if (isCreateDatabase)
                        frmDatabaseSet.intMode = 1;

                    frmDatabaseSet.ShowDialog(this);
                    if (frmDatabaseSet.strConn != "")
                    {
                        strConn = frmDatabaseSet.strConn;
                        //初始化窗口
                        sqlConn.ConnectionString = strConn;


                        sqlComm.CommandText = "SELECT   ID, [num of week], [num of pass], range, [string TAG], [SAP Control Number], [SAP order], [Sales plan], [CW IMEI], [Master data] FROM parameters";
                        sqlConn.Open();
                        sqldr = sqlComm.ExecuteReader();
                        if (sqldr.HasRows)
                        {
                            sqldr.Read();
                            iNumofWeek = int.Parse(sqldr.GetValue(1).ToString());
                            iNumofPass = int.Parse(sqldr.GetValue(2).ToString());
                            iRange = int.Parse(sqldr.GetValue(3).ToString());

                            sTAG = sqldr.GetValue(4).ToString(); ;
                            iSAPControlNumber = int.Parse(sqldr.GetValue(5).ToString()); ;
                            sSAPorder = sqldr.GetValue(6).ToString();
                            sSalesplan = sqldr.GetValue(7).ToString();
                            sCWIMEI = sqldr.GetValue(8).ToString();
                            sMaster = sqldr.GetValue(9).ToString();

                            sqldr.Close();

                        }
                        sqlConn.Close();
                    }
                    else
                    {
                        //this.Close();
                        return;
                    }
                }
                finally
                {
                    sqlConn.Close();
                }

            }
            else  //不存在文件
            {

                /*
                formDatabaseSet frmDatabaseSet = new formDatabaseSet();
                frmDatabaseSet.ShowDialog(this);
                if (frmDatabaseSet.strConn != "")
                {
                    strConn = frmDatabaseSet.strConn;
                    //初始化窗口
                    sqlConn.ConnectionString = strConn;

                    sqlComm.CommandText = "SELECT   ID, [num of week], [num of pass], range, [string TAG], [SAP Control Number], [SAP order], [Sales plan], [CW IMEI], [Master data] FROM parameters";
                    sqlConn.Open();
                    sqldr = sqlComm.ExecuteReader();
                    if (sqldr.HasRows)
                    {
                        sqldr.Read();
                        iNumofWeek = int.Parse(sqldr.GetValue(1).ToString());
                        iNumofPass = int.Parse(sqldr.GetValue(2).ToString());
                        iRange = int.Parse(sqldr.GetValue(3).ToString());

                        sTAG = sqldr.GetValue(4).ToString(); ;
                        iSAPControlNumber = int.Parse(sqldr.GetValue(5).ToString()); ;
                        sSAPorder = sqldr.GetValue(6).ToString();
                        sSalesplan = sqldr.GetValue(7).ToString();
                        sCWIMEI = sqldr.GetValue(8).ToString();
                        sMaster = sqldr.GetValue(9).ToString();
                        sqldr.Close();

                    }
                    sqlConn.Close();
                }
                else
                {
                    //this.Close();
                    return;
                }
                */

                strConn = "";
                return;

            }





        }

        private void produceToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (strConn == "")
            {
                MessageBox.Show("not connect database yet,you can try 'import data offline' now", "Information", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            // 创建此子窗体的一个新实例。
            FormProducts childFormProducts = new FormProducts();
            // 在显示该窗体前使其成为此 MDI 窗体的子窗体。
            childFormProducts.MdiParent = this;

            childFormProducts.strConn = strConn;
            //childFormProducts.WindowState = FormWindowState.Maximized;
            childFormProducts.Show();
        }

        private void optionsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //if (strConn == "")
            //{
            //    MessageBox.Show("not connect database yet,you can try 'import data offline' now", "Information", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            //    return;
            //}
            
            FormParameter formParameter = new FormParameter();
            formParameter.strConn = strConn;
            formParameter.numericUpDownPass.Value = iNumofPass;
            formParameter.numericUpDownWeek.Value = iNumofWeek;

            if (formParameter.ShowDialog() == DialogResult.OK)
            {
                iNumofPass =  (int)formParameter.numericUpDownPass.Value;
                iNumofWeek = (int)formParameter.numericUpDownWeek.Value;
                iRange = int.Parse(Math.Pow(10, (double)formParameter.numericUpDownTAC.Value).ToString());
            }

        }

        private void indentorIToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (strConn == "")
            {
                MessageBox.Show("not connect database yet,you can try 'import data offline' now", "Information", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            
            // 创建此子窗体的一个新实例。
            FormIndentors childFormIndentors = new FormIndentors();
            // 在显示该窗体前使其成为此 MDI 窗体的子窗体。
            childFormIndentors.MdiParent = this;

            childFormIndentors.strConn = strConn;
            //childFormProducts.WindowState = FormWindowState.Maximized;
            childFormIndentors.Show();
        }

        private void preCodeRToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (strConn == "")
            {
                MessageBox.Show("not connect database yet,you can try 'import data offline' now", "Information", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            
            // 创建此子窗体的一个新实例。
            FormTAC childFormTAC = new FormTAC();
            // 在显示该窗体前使其成为此 MDI 窗体的子窗体。
            childFormTAC.MdiParent = this;


            childFormTAC.strConn = strConn;
            childFormTAC.iRange = iRange;
            //childFormProducts.WindowState = FormWindowState.Maximized;
            childFormTAC.Show();
        }

        private void acquireAToolStripMenuItem_Click(object sender, EventArgs e)
        {
             if (strConn == "")
            {
                MessageBox.Show("not connect database yet,you can try 'import data offline' now", "Information", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            // 创建此子窗体的一个新实例。
            FormAcquire childFormAcquire = new FormAcquire();
            // 在显示该窗体前使其成为此 MDI 窗体的子窗体。
            childFormAcquire.MdiParent = this;

            childFormAcquire.strConn = strConn;
            childFormAcquire.iRange = iRange;
            //childFormProducts.WindowState = FormWindowState.Maximized;
            childFormAcquire.Show();
        }

        private void manageToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (strConn == "")
            {
                MessageBox.Show("not connect database yet,you can try 'import data offline' now", "Information", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            
            // 创建此子窗体的一个新实例。
            FormManage childFormManage = new FormManage();
            // 在显示该窗体前使其成为此 MDI 窗体的子窗体。
            childFormManage.MdiParent = this;

            childFormManage.strConn = strConn;
            childFormManage.iNumofPass = iNumofPass;
            childFormManage.iNumofWeek = iNumofWeek;

            childFormManage.iRange = iRange;
            //childFormProducts.WindowState = FormWindowState.Maximized;
            childFormManage.Show();
        }

        private void manufactureMToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (strConn == "")
            {
                MessageBox.Show("not connect database yet,you can try 'import data offline' now", "Information", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            // 创建此子窗体的一个新实例。
            FormManufacture childFormManufacture = new FormManufacture();
            // 在显示该窗体前使其成为此 MDI 窗体的子窗体。
            childFormManufacture.MdiParent = this;

            childFormManufacture.strConn = strConn;
            childFormManufacture.iRange = iRange;
            //childFormProducts.WindowState = FormWindowState.Maximized;
            childFormManufacture.Show();
        }

        private void toolStripButtonAcquire_Click(object sender, EventArgs e)
        {
            acquireAToolStripMenuItem_Click(null, null);
        }

        private void toolStripButtonManufacture_Click(object sender, EventArgs e)
        {
            manufactureMToolStripMenuItem_Click(null, null);
        }

        private void toolStripButtonManage_Click(object sender, EventArgs e)
        {
            manageToolStripMenuItem_Click(null, null);
        }

        private void dataBaseDToolStripMenuItem_Click(object sender, EventArgs e)
        {
            formDatabaseSet frmDatabaseSet = new formDatabaseSet();
            frmDatabaseSet.intMode = 0;//测试模式

            if (frmDatabaseSet.ShowDialog(this) == System.Windows.Forms.DialogResult.OK)
            {
                strConn = frmDatabaseSet.strConn;
                if (strConn == "") //连接错误
                    return;

                //初始化窗口
                sqlConn.ConnectionString = strConn;
                sqlComm.Connection = sqlConn;
                sqlDA.SelectCommand = sqlComm;


            }

        }

        private void importDataIToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (strConn == "")
            {
                MessageBox.Show("not connect database yet,you can try 'import data offline' now", "Information", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            
            // 创建此子窗体的一个新实例。
            FormImportData formImportData = new FormImportData();
            // 在显示该窗体前使其成为此 MDI 窗体的子窗体。
            formImportData.strConn = strConn;
            formImportData.iRange = iRange;
            formImportData.sTAG = sTAG;
            formImportData.iSAPControlNumber = iSAPControlNumber;
            formImportData.sSAPorder = sSAPorder;
            formImportData.sSalesplan = sSalesplan;
            formImportData.sCWIMEI = sCWIMEI;
            formImportData.sMaster = sMaster;

            //childFormProducts.WindowState = FormWindowState.Maximized;
            formImportData.ShowDialog();
        }

        private void toolStripButtonImport_Click(object sender, EventArgs e)
        {
            importDataIToolStripMenuItem_Click(null, null);
        }

        private void createDatabaseCToolStripMenuItem_Click(object sender, EventArgs e)
        {
            formDatabaseSet frmDatabaseSet = new formDatabaseSet();
            frmDatabaseSet.intMode = 1;//

            if (frmDatabaseSet.ShowDialog(this) == System.Windows.Forms.DialogResult.OK)
            {
                strConn = frmDatabaseSet.strConn;
                if (strConn == "") //连接错误
                    return;

                //初始化窗口
                sqlConn.ConnectionString = strConn;
                sqlComm.Connection = sqlConn;
                sqlDA.SelectCommand = sqlComm;


            }
        }

        private void offToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // 创建此子窗体的一个新实例。
            FormImportDataOffLine formImportDataOffLine = new FormImportDataOffLine();
            // 在显示该窗体前使其成为此 MDI 窗体的子窗体。
            formImportDataOffLine.f = this;
            formImportDataOffLine.MdiParent = this;
            formImportDataOffLine.iNumofPass = iNumofPass;
            formImportDataOffLine.iNumofWeek = iNumofWeek;
            formImportDataOffLine.iRange = iRange;
            //childFormProducts.WindowState = FormWindowState.Maximized;
            formImportDataOffLine.Show();
        }

        private void toolStripButtonOffline_Click(object sender, EventArgs e)
        {
            offToolStripMenuItem_Click(null, null);
        }




    }
}
