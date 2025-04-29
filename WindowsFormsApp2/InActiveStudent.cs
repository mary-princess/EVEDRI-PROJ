using Spire.Xls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using WindowsFormsApp2.Class;

namespace WindowsFormsApp2
{
    public partial class InActiveStudent : Form
    {
        public InActiveStudent()
        {
            InitializeComponent();

        }
        Logs myLogs;
        public InActiveStudent(Logs mylogs)
        {
            InitializeComponent();
            this.myLogs = mylogs;
            showStatus("0");
            picProfile.Image = mylogs.Profile;
            picProfile.SizeMode = PictureBoxSizeMode.StretchImage;
            //if (!string.IsNullOrEmpty(myLogs.ImagePath) && File.Exists(myLogs.ImagePath))
            //{
            //    picProfile.Image = Image.FromFile(myLogs.ImagePath);
            //    picProfile.SizeMode = PictureBoxSizeMode.StretchImage;
            //}
        }

        private void Inactive()
        {
            Workbook book = new Workbook();
            book.LoadFromFile(@"C:\Users\User\OneDrive\Desktop\Book1.xlsx");

            Worksheet sheet = book.Worksheets[0];
            DataTable dt = sheet.ExportDataTable();

            dgvInactiveData.DataSource = dt;
        }

        private void btnActiveStatus_Click(object sender, EventArgs e)
        {
            Form2 activeStudent = new Form2(myLogs);
            myLogs.insertLogs(myLogs.GlobalUser, "Visited Active Student");
            activeStudent.Show();
            this.Hide();
        }

        private void btnInactiveStatus_Click(object sender, EventArgs e)
        {
            //InActiveStudent InactiveStudent = new InActiveStudent(myLogs);
            //myLogs.insertLogs(myLogs.GlobalUser, "Visited Inactive Student");
            //InactiveStudent.Show();
            //this.Hide();
        }
        public void showStatus(string status)
        {
            Workbook workbook = new Workbook();
            workbook.LoadFromFile(myLogs.FilePath);
            //workbook.LoadFromFile(@"C:\Users\User\OneDrive\Desktop\Book1.xlsx");

            Worksheet worksheet = workbook.Worksheets[0];
            DataTable dt = worksheet.ExportDataTable();
            DataRow[] dataRow = dt.Select("Status = " + status);

            foreach (DataRow r in dataRow)
            {
                dgvInactiveData.Rows.Add
                    (r[0].ToString(), r[1].ToString(), r[2].ToString(), r[3].ToString(), r[4].ToString(),
                    r[5].ToString(), r[6].ToString(), r[7].ToString(), r[8].ToString(), r[9].ToString(), r[10].ToString()
                   );
            }
        }
        private void InActiveStudent_Load(object sender, EventArgs e)
        {
            lblTime.Text = DateTime.Now.ToString("hh:mm:ss tt");
            lblDate.Text = DateTime.Now.ToString("MM/dd/yyyy");
            lblName.Text = myLogs.GlobalUser;

         
        }

        private void btnDashboard_Click(object sender, EventArgs e)
        {
            Dashboard dashboard = new Dashboard(myLogs);
            myLogs.insertLogs(myLogs.GlobalUser, "Visited Dashboard");
            dashboard.Show();
            this.Hide() ;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            addStudent addStudent = new addStudent(myLogs);
            myLogs.insertLogs(myLogs.GlobalUser, "Visited Add Student");
            addStudent.Show();
            this.Hide() ;
        }

        private void btnLogs_Click(object sender, EventArgs e)
        {
            frmLogs frmLogs = new frmLogs(myLogs);
            myLogs.insertLogs(myLogs.GlobalUser, "Visited Logs");
            frmLogs.Show();
            this.Hide() ;
        }

        private void dgvInactiveData_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void btnSearch_Click(object sender, EventArgs e)
        {

            string searchValue = txtSearch.Text;
            dgvInactiveData.ClearSelection();
            foreach (DataGridViewRow row in dgvInactiveData.Rows)
            {
                if (row.Cells[1].Value.ToString().Equals(searchValue))
                {
                    row.Selected = true;
                    break;
                }

            }
        }
    }
}
