using System;
using System.Collections.Generic;
using System.Text;

namespace DataAccess
{
    public class Utility
    {
        public static string GetDate()
        {
            return string.Format("{0:yyyyMMdd}", DateTime.Now);

        }
        public static string GetTime()
        {
            return string.Format("{0:hhmmss}", DateTime.Now);

        }


        public static string GenLot()
        {
            String lotno;
            lotno = String.Format("{0:yyMMddHHmm}", DateTime.Now);
            return lotno;
        }

        public static string GetExcelColumnName(int columnNumber)
        {
            int dividend  = columnNumber;
            string columnName = "";
            int modulo ;
                while(dividend > 0)
                {
                    modulo = (dividend - 1) % 26;
                    columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                    dividend = Convert.ToInt16((dividend - modulo) / 26);
                }
            
            return columnName;
        }

       public static string AddSingleQuoat(object str)
       {
           string stri;
           stri = "'" + str.ToString() + "'";
           return stri.Trim();
       }
         

        public static string GetMonthName(int month)
        {
            string name;
            month = month % 12==0 ? 12 : month%12;
            switch(month)
            {
                case 1 : name = "January";
                    break;
                case 2: name = "February";
                    break;
                case 3: name = "March";
                    break;
                case 4: name = "April";
                    break;
                case 5: name = "May";
                    break;
                case 6: name = "June";
                    break;
                case 7: name = "July";
                    break;
                case 8: name = "August";
                    break;
                case 9: name = "September";
                    break;
                case 10: name = "October";
                    break;
                case 11: name = "November";
                    break;
                case 12: name = "December";
                    break;

                default: name = "";
                    break;
            }
            return name;
        }
       
    }
}
