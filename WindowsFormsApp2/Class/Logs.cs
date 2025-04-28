using Spire.Xls;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.WebSockets;
using System.Windows.Forms;

namespace WindowsFormsApp2.Class
{
    public class Logs : Usercs
    {

        Workbook book = new Workbook();
        

        public void insertLogs(string user, string message) 
        {
            book.LoadFromFile(@"C:\Users\User\OneDrive\Desktop\Book1.xlsx");
            //book.LoadFromFile(@"C:\Users\ACT-STUDENT\Desktop\Book1.xlsx");
            Worksheet sh = book.Worksheets[1];
            int row = sh.Rows.Length + 1;
            sh.Range[row, 1].Value = user;
            sh.Range[row, 2].Value = message;
            sh.Range[row, 3].Value = DateTime.Now.ToString("MM/dd/yyyy");
            sh.Range[row, 4].Value = DateTime.Now.ToString("hh:mm:ss:tt");
         
            book.SaveToFile(@"C:\Users\User\OneDrive\Desktop\Book1.xlsx", ExcelVersion.Version2016);

        }

        public void showLogs(DataGridView v)
        {
            book.LoadFromFile(@"C:\Users\User\OneDrive\Desktop\Book1.xlsx");
            //book.LoadFromFile(@"C:\Users\ACT-STUDENT\Desktop\Book1.xlsx");

            Worksheet sh = book.Worksheets[1];
            DataTable dt = sh.ExportDataTable();
            v.DataSource = dt;  
        }

     
    }
}
