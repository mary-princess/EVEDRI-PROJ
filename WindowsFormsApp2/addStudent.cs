using Spire.Xls;
using Spire.Xls.Core.Interfaces;
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
    public partial class addStudent : Form
    {
        public addStudent()
        {
            InitializeComponent();
        }
        Logs myLogs;

        public addStudent(Logs mylogs)
        {
            InitializeComponent();
            this.myLogs = mylogs;

            picProfile.Image = mylogs.Profile;
            picProfile.SizeMode = PictureBoxSizeMode.StretchImage;
            //if (!string.IsNullOrEmpty(myLogs.ImagePath) && File.Exists(myLogs.ImagePath))
            //{
            //    picProfile.Image = Image.FromFile(myLogs.ImagePath);
            //    picProfile.SizeMode = PictureBoxSizeMode.StretchImage;
            //}
        }
        private void btnInsert_Click(object sender, EventArgs e)
        {
            Workbook book = new Workbook();
            book.LoadFromFile(@"C:\Users\User\OneDrive\Desktop\Book1.xlsx");
            //book.LoadFromFile(@"C:\Users\ACT-STUDENT\Desktop\Book1.xlsx");
            Worksheet sheet = book.Worksheets[0];

            bool userExists = false;
            int totalRows = sheet.Rows.Length;

            for(int i = 2; i <= totalRows; i++)
            {
                string existsUser = sheet.Range[i, 8].Value;
                string existsPass = sheet.Range[i, 9].Value;
                
                if(existsUser == txtUsername.Text || existsPass == txtPassword.Text)
                {
                    userExists = true;
                    break;
                }
            }
            if (userExists)
            {
                MessageBox.Show("Username and Password already exist! Please provide another one", "Duplication Found", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }

            bool errorRestriction = false;

            error.Clear();

            if (!string.IsNullOrEmpty(txtName.Text))
            {
                error.SetError(txtName, "Name is required.");
                errorRestriction = true;
            }
            if(!radMale.Checked && !radFemale.Checked)
            {
                error.SetError(label2, "Select Gender.");
                errorRestriction = true;
            }
            if(!chkBadminton.Checked && !chkBball.Checked && !chkVball.Checked)
            {
                error.SetError(label3, "Select at least one hobby.");
                errorRestriction = true;
            }
            if(cboDegree.SelectedItem == null)
            {
                error.SetError(cboDegree, "Select a degree.");
                errorRestriction = true;
            }
            if(cboColor.SelectedItem == null)
            {
                error.SetError(cboColor, "Select a color.");
                errorRestriction = true;
            }
            if (!string.IsNullOrEmpty(txtSayings.Text))
            {
                error.SetError(txtSayings, "Sayings is required.");
                errorRestriction = true;
            }
            if(dtpBirthdate.Value.Date > DateTime.Today)
            {
                error.SetError(dtpBirthdate, "Birthdate is required.");
                error.SetError(lblAge, "Fill in the birthdate field.");
                errorRestriction = true;
            }
            if (!string.IsNullOrEmpty(txtUsername.Text))
            {
                error.SetError(txtUsername, "Username is required.");
                errorRestriction = true;
            }
            if (!string.IsNullOrEmpty(txtPassword.Text))
            {
                error.SetError(txtPassword, "Password is required.");
                errorRestriction = true;
            }
            if (!myLogs.ValidateEmailAddress(txtEmailAddress.Text) || !string.IsNullOrEmpty(txtEmailAddress.Text))
            {
                error.SetError(txtEmailAddress, "Invalid Email Address");
                errorRestriction = true;
            }
            if (errorRestriction)
            {
                MessageBox.Show("Please correct the highlighted errors before proceeding", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
           

            string name = txtName.Text, gender = "", hobbies = ""; 

            if (radMale.Checked)
            {
                gender = "Male";
            }
            else if (radFemale.Checked)
            {
                gender = "Female";
            }

            if (chkBadminton.Checked) hobbies += chkBadminton.Text + ", ";
            if (chkBball.Checked) hobbies += chkBball.Text + ", ";
            if (chkVball.Checked) hobbies += chkVball.Text + ", ";
            hobbies = hobbies.TrimEnd(',', ' ');
            
            string degree = cboDegree.SelectedItem.ToString();
            string color = cboColor.SelectedItem.ToString();
            string sayings = txtSayings.Text;
            string username = txtUsername.Text;
            string password = txtPassword.Text; 
            string email = txtEmailAddress.Text;
            string imagePath = txtProfile.Text;


            int age = myLogs.CalculateAge(dtpBirthdate.Value);
            lblAge.Text = age.ToString();

            int row = sheet.LastRow + 1;

            sheet.Range[row, 1].Value = name;
            sheet.Range[row, 2].Value = gender;
            sheet.Range[row, 3].Value = hobbies;
            sheet.Range[row, 4].Value = degree;
            sheet.Range[row, 5].Value = color;
            sheet.Range[row, 6].Value = sayings;
            sheet.Range[row, 7].Value = age.ToString();
            sheet.Range[row, 8].Value = username;
            sheet.Range[row, 9].Value = password;
            sheet.Range[row, 10].Value = "1";
            sheet.Range[row, 11].Value = imagePath;
            sheet.Range[row, 12].Value = email;

            book.SaveToFile(@"C:\Users\User\OneDrive\Desktop\Book1.xlsx", ExcelVersion.Version2016);
            //book.SaveToFile(@"C:\Users\ACT-STUDENT\Desktop\Book1.xlsx", ExcelVersion.Version2016);

            myLogs.insertLogs(myLogs.GlobalUser, "Inserting Info");
            DataTable dt = sheet.ExportDataTable();
            Form2 form2 = new Form2();
            form2.dgvData.DataSource = dt;
            MessageBox.Show("Successfully added data.", "");
        }

        private void btnDashboard_Click(object sender, EventArgs e)
        {
            Dashboard dashboard = new Dashboard(myLogs);
            myLogs.insertLogs(myLogs.GlobalUser, "Visited Dashbooard");
            dashboard.Show();
            this.Hide();
        }

        private void btnActiveStatus_Click(object sender, EventArgs e)
        {
            Form2 form2 = new Form2(myLogs);
            myLogs.insertLogs(myLogs.GlobalUser, "Visited Active Student");
            form2.Show();
            this.Hide();
        }

        private void btnInactiveStatus_Click(object sender, EventArgs e)
        {
            InActiveStudent inActiveStudent = new InActiveStudent(myLogs);
            myLogs.insertLogs(myLogs.GlobalUser, "Visited Inactive Student");
            inActiveStudent.Show();
            this.Hide();
        }

        private void btnLogs_Click(object sender, EventArgs e)
        {
            frmLogs logs = new frmLogs(myLogs);
            myLogs.insertLogs(myLogs.GlobalUser, "Visited Logs");
            logs.Show();
            this.Hide();
        }

        private void btnLogout_Click(object sender, EventArgs e)
        {
            Login login = new Login(myLogs);
            myLogs.insertLogs(myLogs.GlobalUser, "Log Out");
            login.Show();
            this.Hide();
        }

        private void addStudent_Load(object sender, EventArgs e)
        {
            lblTime.Text = DateTime.Now.ToLongTimeString();
            lblDate.Text = DateTime.Now.ToLongDateString();
            lblName.Text = myLogs.GlobalUser;
        }

        private void btnProfile_Click(object sender, EventArgs e)
        {
            //Workbook book = new Workbook();
            //book.LoadFromFile(@"C:\Users\User\OneDrive\Desktop\Book1.xlsx");
            ////book.LoadFromFile(@"C:\Users\ACT-STUDENT\Desktop\Book1.xlsx");

            //Worksheet sheet = book.Worksheets[0];

            OpenFileDialog file = new OpenFileDialog();
            file.Filter = "Image Files|*.jpg;*.jpeg;*.png;*.bmp;*.gif";


            if (file.ShowDialog() == DialogResult.OK) 
            {
                txtProfile.Text = file.FileName;


            }

        }


        //TO BE CONTINUED
        //public void showStatus(string status)
        //{
        //    Workbook book = new Workbook();
        //    book.LoadFromFile(@"C:\Users\ACT-STUDENT\Desktop\Book1.xlsx");

        //    Worksheet sh = book.Worksheets[0];
        //    DataTable dt = sh.ExportDataTable();
        //    DataRow[] row = dt.Select("Status = " + status);

        //    foreach (DataRow row2 in row)
        //    {
        //    }
        //}
    }
}
