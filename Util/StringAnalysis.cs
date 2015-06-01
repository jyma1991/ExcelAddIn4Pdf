using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelAddIn4Pdf.Util
{
    class StringAnalysis
    {
        /// <summary>判断日期</summary>
        /// <param name="strDate">须判断字符串</param>
        /// <returns>是否为日期</returns>
        public bool IsDate(string strDate)
        {
            try
            {
                DateTime.Parse(strDate);
                return true;
            }
            catch
            {
                return false;
            }
        }

        public DateTime GetDays(string sOneDay)
        {
            string sDateTime = sOneDay.Substring(0, sOneDay.IndexOf("&"));
            return Convert.ToDateTime(sDateTime); 
        }

        public SortedList<DateTime, string> GetTravel(DateTime dtDay, string sOneDay)
        {
            SortedList<DateTime, string> dayTravel = new SortedList<DateTime, string>();
            string[] sLine = null;
            if (sOneDay.Contains("#"))sLine = sOneDay.Split('#');
            if(sLine!=null)
                foreach (var item in sLine)
                {
                    if (item.Contains("$")&& item.Contains("@"))
                    {
                        int i = item.IndexOf('$');
                        int j = item.IndexOf('@');
                        DateTime dt = dtDay.Date + Convert.ToDateTime(item.Substring(0, item.IndexOf('$'))).TimeOfDay;
                        if (dayTravel.ContainsKey(dt))
                            dayTravel[dt] += item.Substring(i + 1,j-i-1);
                        else
                            dayTravel.Add(dt, item.Substring(i + 1, j - i - 1));
                    }
                    else if(item.Contains("$"))
                    {
                        int i = item.IndexOf('$');
                        DateTime dt = dtDay.Date+Convert.ToDateTime(item.Substring(0, item.IndexOf('$'))).TimeOfDay;
                        if (dayTravel.ContainsKey(dt)) 
                            dayTravel[dt] += item.Substring(i + 1);
                        else
                            dayTravel.Add(dt, item.Substring(i + 1));
                    }
                }
            return dayTravel;
        }
    }
}
