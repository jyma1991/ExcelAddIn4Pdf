using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelAddIn4Pdf.Util
{
    class CairnsTravel
    {
        public string DayTitle { get; set; }

        public SortedList<DateTime, string> DayTravel { get; set; }

        public CairnsTravel(string Title,SortedList<DateTime, string> travel)
        {
            this.DayTitle = Title;
            this.DayTravel = travel;
        }
    }
}
