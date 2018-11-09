using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Windows.Forms;

namespace imesManger
{
    public static class ClassIm
    {
        public static string sYear = "";


        public static void ClearDataGridViewErrorText(DataGridView dvIn)
        {
            if (dvIn.CurrentCell == null)
                return;
            for (int i = 0; i < dvIn.ColumnCount; i++)
            {
                dvIn.Rows[dvIn.CurrentCell.RowIndex].Cells[i].ErrorText = String.Empty;
            }
        }

        #region   得到当前日期是该年度的第几周，此处需要注意的是：不能从该年的1-1算起，判断该年的1-1是星期几，如果是周一，那么从1-1算起，如果非周一，要将偏移量减去
        ///   <summary > 
        ///   能够得到该年有多少周，传递参数的时候“当前年+12-31” 
        ///   </summary > 
        ///   <param   name="dt" > </param > 
        ///   <returns > </returns > 
        /*
        public static string GetWeek(DateTime dt)
        {
            string returnStr = "";
            //System.DateTime fDt = DateTime.Parse(dt.Year.ToString() + "-01-01");
            System.DateTime fDt = new DateTime(dt.Year,1,1);
            int k = Convert.ToInt32(fDt.DayOfWeek);//得到该年的第一天是周几 
            if (k == 0)
            {
                k = 7;
            }
            int l = Convert.ToInt32(dt.DayOfYear);//得到当天是该年的第几天 
            l = l - (7 - k + 1);
            if (l <= 0)
            {
                returnStr = "1";
            }
            else
            {
                if (l % 7 == 0)
                {
                    returnStr = (l / 7 + 1).ToString();
                }
                else
                {
                    returnStr = (l / 7 + 2).ToString();//不能整除的时候要加上前面的一周和后面的一周 
                }
            }
            return returnStr;
        }
        */

        public static int GetWeek(DateTime dt)
        {
            System.DateTime fDt = new DateTime(dt.Year, 1, 1); 
            int k = Convert.ToInt32(fDt.DayOfWeek);//得到该年的第一天是周几 

            //得到1月1日周六
            int weeknow = Convert.ToInt32(fDt.DayOfWeek);
            if (weeknow == 0)
                weeknow = 7;

            //如果为周六，日 向前推0，1天
            weeknow = (weeknow >= 6 ? (weeknow-6) : (weeknow + 1));
            DateTime dt0 = fDt.AddDays(-1*weeknow);

            TimeSpan ts = dt - dt0;
            sYear = dt.Year.ToString();
            return ((int)(ts.Days / 7)+1);
            
        }

        public static string GetWeek(DateTime dt,int iDifference)
        {
            string returnStr = "";

            //得到
            int iWeek = GetWeek(dt);
            int iTemp = iWeek + iDifference;

            DateTime dt1 = new DateTime(dt.Year,12,31);
            DateTime dt2 = new DateTime(dt.Year-1,12,31);

            if (iTemp > 0) //向后算
            {
                int iWeek1=GetWeek(dt1);

                if (iTemp > iWeek1)
                {
                    sYear=(dt.Year+1).ToString();
                    returnStr = (iTemp - iWeek1).ToString();
                }
                else
                {
                    sYear=dt.Year.ToString();
                    returnStr=iTemp.ToString();
                }
            }
            else //向前算
            {
                int iWeek2 = GetWeek(dt2);
                sYear = (dt.Year - 1).ToString();
                returnStr = (iWeek2+iTemp).ToString();
            }

            return returnStr;
        }

        public static string GetWeek(int iYear, int iWeek, int iDifference)
        {
            string returnStr = "";


            //得到
            int iTemp = iWeek + iDifference;

            DateTime dt1 = new DateTime(iYear, 12, 31);
            DateTime dt2 = new DateTime(iYear - 1, 12, 31);



            if (iTemp > 0) //向后算
            {
                int iWeek1 = GetWeek(dt1);

                if (iTemp > iWeek1)
                {
                    sYear = (iYear + 1).ToString();
                    returnStr = (iTemp - iWeek1).ToString();
                }
                else
                {
                    sYear = iYear.ToString();
                    returnStr = iTemp.ToString();
                }
            }
            else //向前算
            {
                int iWeek2 = GetWeek(dt2);
                sYear = (iYear - 1).ToString();
                returnStr = (iWeek2 + iTemp).ToString();
            }

            return returnStr;
        }

        public static DateTime GetDate(int iYear, int iWeek) //返回周第一天
        {
            //第一周
            DateTime dt1 = new DateTime(iYear, 1, 1);
            if (iWeek == 1)
            {
                return dt1;
            }

            //最后一周
            DateTime dt2 = new DateTime(iYear, 12, 31);
            if (iWeek >= GetWeek(dt2))
            {
                return dt2;
            }


            //星期一为第一天
            int weeknow = Convert.ToInt32(dt1.DayOfWeek);

            //因为是以星期一为第一天，所以要判断weeknow等于0时，要向前推6天。
            weeknow = (weeknow == 0 ? (7 - 1) : (weeknow - 1));
            int daydiff = (-1) * weeknow;

            //1月1号周第一天
            DateTime dt0 = dt1.AddDays(daydiff);
            DateTime dt = dt1.AddDays((iWeek-1)*7-1);

            return dt;


        }
        #endregion 

        public static string GetPCode(string PCode) //去除0
        {
            if (PCode == "")
                return "";

            while(PCode.Length>1)
            {
                if (PCode.Substring(0, 1) == "0") 
                    PCode = PCode.Substring(1, PCode.Length - 1);
                else
                    break;
            }

            return PCode;

        }

    }

    public class ClassIndexofOrder
    {
        public int iFrom = 0;
        public int iTo = 0;
        public int iNo = 0;

        public ClassIndexofOrder(int iFrom, int iTo, int iNo)
        {
            this.iFrom = iFrom; this.iTo = iTo; this.iNo = iNo;
        }

    }


}
