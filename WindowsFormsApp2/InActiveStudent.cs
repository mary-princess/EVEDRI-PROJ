using Spire.Xls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
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
                    r[5].ToString(), r[6].ToString(), r[7].ToString(), r[8].ToString(), r[9].ToString(), r[10].ToString(),
                    r[11].ToString()
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
            this.Hide();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            addStudent addStudent = new addStudent(myLogs);
            myLogs.insertLogs(myLogs.GlobalUser, "Visited Add Student");
            addStudent.Show();
            this.Hide();
        }

        private void btnLogs_Click(object sender, EventArgs e)
        {
            frmLogs frmLogs = new frmLogs(myLogs);
            myLogs.insertLogs(myLogs.GlobalUser, "Visited Logs");
            frmLogs.Show();
            this.Hide();
        }

        private void dgvInactiveData_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void btnSearch_Click(object sender, EventArgs e)
        {
            string searchValue = txtSearch.Text.Trim().ToLower();
            dgvInactiveData.ClearSelection();

            if (string.IsNullOrEmpty(searchValue))
            {
                MessageBox.Show("Please enter a value to search.", "Search", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            dgvInactiveData.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            bool matchFound = false;
            foreach (DataGridViewRow row in dgvInactiveData.Rows)
            {
                if (row.Cells[0].Value.ToString().Equals(searchValue))
                    row.Selected = true;
                else if (matchFound)
                {
                    MessageBox.Show("No matching records found.", "Search", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
        }

        private void dgvInactiveData_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0 && e.RowIndex < dgvInactiveData.Rows.Count)
            {
                int r = e.RowIndex;

                DialogResult result =
                    MessageBox.Show("Are you sure you want to update this data?", "Confirm Update", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                if (result == DialogResult.Yes)
                {
                    Form1 form1 = new Form1(myLogs);
                    form1.Show();


                    form1.lblID.Text = r.ToString();
                    form1.txtName.Text = dgvInactiveData.Rows[r].Cells[0].Value.ToString();
                    string gender = dgvInactiveData.Rows[r].Cells[1].Value.ToString();
                    form1.radMale.Checked = (gender == "Male");
                    form1.radFemale.Checked = (gender == "Female");


                    string hobbies = dgvInactiveData.Rows[r].Cells[2].Value.ToString();
                    string[] arrayHobbies = hobbies.Split(',');

                    form1.chkBadminton.Checked = false;
                    form1.chkBball.Checked = false;
                    form1.chkVball.Checked = false;

                    foreach (string s in arrayHobbies)
                    {
                        string trim = s.Trim();
                        if (trim == "Volleyball") form1.chkVball.Checked = true;
                        if (trim == "Basketball") form1.chkBball.Checked = true;
                        if (trim == "Badminton") form1.chkBadminton.Checked = true;

                    }

                    form1.cboDegree.SelectedItem = dgvInactiveData.Rows[r].Cells[3].Value.ToString();
                    form1.cboColor.SelectedItem = dgvInactiveData.Rows[r].Cells[4].Value.ToString();
                    form1.txtSayings.Text = dgvInactiveData.Rows[r].Cells[5].Value.ToString();

                    DateTime birthDate;
                    if (DateTime.TryParse(dgvInactiveData.Rows[r].Cells[6].Value.ToString(), out birthDate))
                    {
                        form1.dtpBirthdate.Value = birthDate;
                        int age = myLogs.CalculateAge(birthDate);
                        form1.lblAge.Text = age.ToString();
                    }




                    form1.txtUsername.Text = dgvInactiveData.Rows[r].Cells[7].Value.ToString();
                    form1.txtPassword.Text = dgvInactiveData.Rows[r].Cells[8].Value.ToString();
                    //form1.cboStatus.SelectedItem = dgvInactiveData.Rows[r].Cells[9].Value.ToString();
                    string imagePath = dgvInactiveData.Rows[r].Cells[10].Value.ToString();
                    form1.txtEmailAddress.Text = dgvInactiveData.Rows[r].Cells[11].Value.ToString();    
                    if (!string.IsNullOrEmpty(imagePath) && File.Exists(imagePath))
                    {
                        form1.picProfile.Image = Image.FromFile(imagePath);
                        form1.picProfile.SizeMode = PictureBoxSizeMode.StretchImage;
                        form1.txtProfile.Text = imagePath;
                    }
                    else
                    {
                        form1.picProfile.Image = null;
                        form1.txtProfile.Text = "";

                    }

                    
                    this.Hide();
                }
            }
        }

        private void btnRestore_Click(object sender, EventArgs e)
        {
            if (dgvInactiveData.CurrentRow == null)
            {
                MessageBox.Show("Please select a row to restore data.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            DialogResult result = MessageBox.Show("Are you sure you want to restore this data?", "Confirm", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (result == DialogResult.No)
            {
                return;
            }
            Workbook workbook = new Workbook();
            workbook.LoadFromFile(myLogs.FilePath);
            Worksheet sh = workbook.Worksheets[0];

            int totalRows = sh.Rows.Length;

            foreach (DataGridViewRow selectedRow in dgvInactiveData.SelectedRows)
            {
                string restore = selectedRow.Cells[7].Value.ToString();
                for (int i = 2; i <= totalRows; i++)
                {

                    if (sh.Range[i, 8].Value.ToString() == restore)
                    {
                        sh.Range[i, 10].Value = "1";
                        break;
                    }
                }

            }



            workbook.SaveToFile(myLogs.FilePath, ExcelVersion.Version2016);
            myLogs.insertLogs(myLogs.GlobalUser, "Restored an account");


            dgvInactiveData.Rows.Clear();
            showStatus("0");

            MessageBox.Show("Data restored successfully.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void btnLogout_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("Are you sure you want to log out?", "Log Out", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (result == DialogResult.No)
            {
                return;
            }
            Login login = new Login(myLogs);
            myLogs.insertLogs(myLogs.GlobalUser, "Log Out");
            login.Show();
            this.Hide();
        }

        private void pictureBox8_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("Are you sure you want to close this application", "Application Exit", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (result == DialogResult.No)
            {
                return;
            }
            Application.Exit();
        }
    }
}
