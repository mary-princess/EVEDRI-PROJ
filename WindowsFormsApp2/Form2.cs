using Spire.Xls;
using Spire.Xls.Core;
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
using static Guna.UI2.Native.WinApi;

namespace WindowsFormsApp2
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
         
        }
        Logs myLogs;
        public Form2(Logs logs)
        {

            InitializeComponent();
            this.myLogs = logs;
            showStatus("1");
            picProfile.Image = myLogs.Profile;
            picProfile.SizeMode = PictureBoxSizeMode.StretchImage;
            //if (!string.IsNullOrEmpty(myLogs.ImagePath) && File.Exists(myLogs.ImagePath))
            //{
            //    picProfile.Image = Image.FromFile(myLogs.ImagePath);
            //    picProfile.SizeMode = PictureBoxSizeMode.StretchImage;
            //}
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            lblTime.Text = DateTime.Now.ToString("hh:mm:ss:tt");
            lblDate.Text = DateTime.Now.ToString("MM/dd/yyyy");
            lblName.Text = myLogs.GlobalUser;

           
        }
        Login login = new Login();
        public void LoadExcelFile()
        {
            Workbook book = new Workbook();
            book.LoadFromFile(@"C:\Users\User\OneDrive\Desktop\Book1.xlsx");
            //book.LoadFromFile(@"C:\Users\ACT-STUDENT\Desktop\Book1.xlsx");


            Worksheet sheet = book.Worksheets[0];
            DataTable dt = sheet.ExportDataTable();
            dgvData.DataSource = dt;
        }

        //public void insertTable(int id, string name, string gender, string hobbies, string favoriteColor, string sayings, string username, string password)
        //{
        //    dgvData.Rows.Add(id, name, gender, hobbies, favoriteColor, sayings, username, password);

        //    //int r = dgvData.Rows.Add();
        //    //dgvData.Rows[r].Cells[0].Value = id;
        //    //dgvData.Rows[r].Cells[1].Value = name;
        //    //dgvData.Rows[r].Cells[2].Value = gender;
        //    //dgvData.Rows[r].Cells[3].Value = hobbies;
        //    //dgvData.Rows[r].Cells[4].Value = favoriteColor;
        //    //dgvData.Rows[r].Cells[5].Value = sayings;

        //}

        public void showStatus(string status) 
        {
            Workbook workbook = new Workbook();
            workbook.LoadFromFile(@"C:\Users\User\OneDrive\Desktop\Book1.xlsx");

            Worksheet worksheet = workbook.Worksheets[0];
            DataTable dt = worksheet.ExportDataTable();
            DataRow[] dataRow = dt.Select("Status = " + status);

            foreach (DataRow r in dataRow)
            {
                dgvData.Rows.Add
                    (r[0].ToString(), r[1].ToString(), r[2].ToString(), r[3].ToString(), r[4].ToString(),
                    r[5].ToString(), r[6].ToString(), r[7].ToString(), r[8].ToString(), r[9].ToString(), r[10].ToString()
                   );
            }

        }

        private void dgvData_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            //foreach (DataGridViewRow row in this.dgvData.SelectedRows)
            //{
            //    dgvData.Rows.RemoveAt(row.Index);
            //}

            //LATEST
            //if (dgvData.SelectedRows.Count > 0)
            //{
            //    int row= dgvData.CurrentCell.RowIndex;

            //    string name = dgvData.Rows[row].Cells[0].Value.ToString();
            //    string gender = dgvData.Rows[row].Cells[1].Value.ToString();
            //    string hobbies = dgvData.Rows[row].Cells[2].Value.ToString();
            //    string favoriteColor = dgvData.Rows[row].Cells[3].Value.ToString();
            //    string sayings = dgvData.Rows[row].Cells[4].Value.ToString();
            //    string username = dgvData.Rows[row].Cells[5].Value.ToString();
            //    string password = dgvData.Rows[row].Cells[5].Value.ToString();

            //    updateStatus(name, gender, hobbies, favoriteColor, sayings, username, password);

            //    dgvData.Rows.RemoveAt(row);

            //}

            //else
            //{
            //    MessageBox.Show("Please select a student to delete.", "No Row Selected", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //}

            Workbook workbook = new Workbook();
            //workbook.LoadFromFile(@"C:\Users\ACT-STUDENT\Desktop\Book1.xlsx");
            workbook.LoadFromFile(@"C:\Users\User\OneDrive\Desktop\Book1.xlsx");

            Worksheet sh = workbook.Worksheets[0];

            int r = dgvData.CurrentCell.RowIndex + 2;
            sh.Range[r, 10].Value = "0";

            showStatus("1");

            //workbook.SaveToFile(@"C:\Users\ACT-STUDENT\Desktop\Book1.xlsx", ExcelVersion.Version2016);
            workbook.LoadFromFile(@"C:\Users\User\OneDrive\Desktop\Book1.xlsx", ExcelVersion.Version2016);

            DataTable dt = sh.ExportDataTable();
            dgvData.DataSource = dt;



        }

        //private void updateStatus(string name, string gender, string hobbies, string color, string sayings, string username, string password)
        //{
        //    Workbook book = new Workbook();
        //    book.LoadFromFile(@"C:\Users\User\OneDrive\Desktop\Book1.xlsx");
        //    //book.LoadFromFile(@"C:\Users\ACT-STUDENT\Desktop\Book1.xlsx");


        //    Worksheet sheet = book.Worksheets[0];

        //    int row = -1;
        //    for (int i = 1; i <= sheet.Rows.Length; i++)  // Starting from 1 assuming row 0 is headers
        //    {
        //        if (sheet.Range[i, 6].Value.ToString() == username)  // Match username to find the row
        //        {
        //            row = i;
        //            break;
        //        }
        //    }

        //    if (row != -1)
        //    {
        //        // Set the Status column (last column) to 0 (inactive)
        //        sheet.Range[row, 8].Value2 = 0;  // Column 8 should be the Status column

        //        // Save the changes to the Excel file
        //        book.SaveToFile(@"C:\Users\User\OneDrive\Desktop\Book1.xlsx");

        //        MessageBox.Show("Student status updated to inactive.", "Status Updated", MessageBoxButtons.OK, MessageBoxIcon.Information);
        //    }
        //    else
        //    {
        //        MessageBox.Show("Student not found in the Active Students list.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        //    }

        //}

        private void dgvData_MouseDoubleClick(object sender, MouseEventArgs e)
        {
           

           
        }

        private void dgvData_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0  && e.RowIndex < dgvData.Rows.Count)
            {
                int r = e.RowIndex;

                DialogResult result =
                    MessageBox.Show("Are you sure you want to update this data?", "Confirm Update", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                if(result == DialogResult.Yes)
                {
                    Form1 form1 = new Form1(myLogs);
                    form1.Show();

                    
                    form1.lblID.Text = r.ToString();
                    form1.txtName.Text = dgvData.Rows[r].Cells[0].Value.ToString();
                    string gender = dgvData.Rows[r].Cells[1].Value.ToString();
                    form1.radMale.Checked = (gender == "Male");
                    form1.radFemale.Checked = (gender == "Female");


                    string hobbies = dgvData.Rows[r].Cells[2].Value.ToString();
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

                    form1.cboDegree.SelectedItem = dgvData.Rows[r].Cells[3].Value.ToString();
                    form1.cboColor.SelectedItem = dgvData.Rows[r].Cells[4].Value.ToString();
                    form1.txtSayings.Text = dgvData.Rows[r].Cells[5].Value.ToString();

                    DateTime birthDate;
                    if (DateTime.TryParse(dgvData.Rows[r].Cells[6].Value.ToString(), out birthDate))
                    {
                        form1.dtpBirthdate.Value = birthDate;
                        int age = myLogs.CalculateAge(birthDate);
                        form1.lblAge.Text = age.ToString();
                    }
                    

                  

                    form1.txtUsername.Text = dgvData.Rows[r].Cells[7].Value.ToString();
                    form1.txtPassword.Text = dgvData.Rows[r].Cells[8].Value.ToString();

                    string imagePath = dgvData.Rows[r].Cells[10].Value.ToString();

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

                    form1.txtEmailAddress.Text = dgvData.Rows[r].Cells[11].Value.ToString();
                    this.Hide();


                }
            }

            //int r = dgvData.CurrentCell.RowIndex;
            //Form1 form1 = new Form1(myLogs);
            //form1.Show();
            //Form1 form1 = (Form1)Application.OpenForms["Form1"];

            //form1.lblID.Text = r.ToString();
            //form1.txtName.Text = dgvData.Rows[r].Cells[0].Value.ToString();
            //string gender = dgvData.Rows[r].Cells[1].Value.ToString();
            //form1.radMale.Checked = (gender == "Male");
            //form1.radFemale.Checked = (gender == "Female");


            //string hobbies = dgvData.Rows[r].Cells[2].Value.ToString();
            //string[] arrayHobbies = hobbies.Split(',');

            //form1.chkBadminton.Checked = false;
            //form1.chkBball.Checked = false;
            //form1.chkVball.Checked = false;

            //foreach (string s in arrayHobbies)
            //{
            //    string trim = s.Trim();
            //    if (trim == "Volleyball") form1.chkVball.Checked = true;
            //    if (trim == "Basketball") form1.chkBball.Checked = true;
            //    if (trim == "Badminton") form1.chkBadminton.Checked = true;

            //}

            //form1.cboDegree.SelectedItem = dgvData.Rows[r].Cells[3].Value.ToString();
            //form1.cboColor.SelectedItem = dgvData.Rows[r].Cells[4].Value.ToString();
            //form1.txtSayings.Text = dgvData.Rows[r].Cells[5].Value.ToString();

            //DateTime birthDate = DateTime.Parse(dgvData.Rows[r].Cells[6].Value.ToString());
            //form1.dtpBirthdate.Value = birthDate;

            //int age = myLogs.CalculateAge(birthDate);
            //form1.lblAge.Text = age.ToString();

            //form1.txtUsername.Text = dgvData.Rows[r].Cells[7].Value.ToString();
            //form1.txtPassword.Text = dgvData.Rows[r].Cells[8].Value.ToString();

            //string imagePath = dgvData.Rows[r].Cells[10].Value?.ToString();

            //if (!string.IsNullOrEmpty(imagePath) && File.Exists(imagePath))
            //{
            //    form1.picProfile.Image = Image.FromFile(imagePath);
            //    form1.picProfile.SizeMode = PictureBoxSizeMode.StretchImage;
            //    form1.txtProfile.Text = imagePath;
            //}
            //else
            //{
            //    form1.picProfile.Image = null;
            //    form1.txtProfile.Text = "";

            //}

            //this.Hide();


        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Hide();
        }

        private void btnSearch_Click(object sender, EventArgs e)
        {
           

            string searchValue = txtSearch.Text;
            dgvData.ClearSelection();
            foreach(DataGridViewRow row in dgvData.Rows) 
            {
                if (row.Cells[1].Value.ToString().Equals(searchValue))
                {
                    row.Selected = true;
                    break;
                }

            }

        }

        private void btnInactiveStatus_Click(object sender, EventArgs e)
        {
            InActiveStudent inActiveStudent = new InActiveStudent(myLogs);
            myLogs.insertLogs(myLogs.GlobalUser, "Visited Inactive Student");

            inActiveStudent.Show();
            this.Hide();
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
            frmLogs log = new frmLogs(myLogs);
            myLogs.insertLogs(myLogs.GlobalUser, "Visited Logs");
            log.Show();
            this.Hide();

        }
    }
}
