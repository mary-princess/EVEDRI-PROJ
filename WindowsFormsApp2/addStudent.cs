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
          
        }
        private void btnInsert_Click(object sender, EventArgs e)
        {
            Workbook book = new Workbook();
            book.LoadFromFile(myLogs.FilePath);
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
                    
                }
            }
            if (userExists)
            {
                MessageBox.Show("Username and Password already exist. Please provide another one.", "Duplication Found", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }

            bool errorRestriction = false;

            error.Clear();

            if (string.IsNullOrEmpty(txtName.Text))
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
            if (string.IsNullOrEmpty(txtSayings.Text))
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
            if (string.IsNullOrEmpty(txtUsername.Text))
            {
                error.SetError(txtUsername, "Username is required.");
                errorRestriction = true;
            }
            if (string.IsNullOrEmpty(txtPassword.Text))
            {
                error.SetError(txtPassword, "Password is required.");
                errorRestriction = true;
            }
            if (cboStatus.SelectedItem == null)
            {
                error.SetError(cboStatus, "Select a status.");
                errorRestriction = true;
            }
            if (!myLogs.ValidateEmailAddress(txtEmailAddress.Text) || string.IsNullOrEmpty(txtEmailAddress.Text))
            {
                error.SetError(txtEmailAddress, "Invalid Email Address");
                errorRestriction = true;
            }
            if (string.IsNullOrEmpty(txtProfile.Text))
            {
                error.SetError(txtProfile, "Profile is required.");
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
            string status = cboStatus.SelectedItem.ToString();
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
            sheet.Range[row, 10].Value = status;
            sheet.Range[row, 11].Value = imagePath;
            sheet.Range[row, 12].Value = email;

            book.SaveToFile(myLogs.FilePath, ExcelVersion.Version2016);

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

        private void addStudent_Load(object sender, EventArgs e)
        {
            lblTime.Text = DateTime.Now.ToLongTimeString();
            lblDate.Text = DateTime.Now.ToLongDateString();
            lblName.Text = myLogs.GlobalUser;
        }

        private void btnProfile_Click(object sender, EventArgs e)
        {
           

            OpenFileDialog file = new OpenFileDialog();
            file.Filter = "Image Files|*.jpg;*.jpeg;*.png;*.bmp;*.gif";


            if (file.ShowDialog() == DialogResult.OK) 
            {
                txtProfile.Text = file.FileName;


            }

        }

        private void dtpBirthdate_ValueChanged(object sender, EventArgs e)
        {
            int age = myLogs.CalculateAge(dtpBirthdate.Value);
            lblAge.Text = age.ToString();
        }

        private void dtpBirthdate_MouseClick(object sender, MouseEventArgs e)
        {
            int age = myLogs.CalculateAge(dtpBirthdate.Value);
            lblAge.Text = age.ToString();
        }

        private void txtName_MouseClick(object sender, MouseEventArgs e)
        {
            error.SetError(txtName, string.Empty);
        }

        private void cboColor_SelectedIndexChanged(object sender, EventArgs e)
        {
           
        }

        private void cboColor_MouseClick(object sender, MouseEventArgs e)
        {
            error.SetError(cboColor, string.Empty);

        }

        private void cboDegree_MouseClick(object sender, MouseEventArgs e)
        {
            error.SetError(cboDegree, string.Empty);
        }

        private void txtSayings_MouseClick(object sender, MouseEventArgs e)
        {
            error.SetError(txtSayings, string.Empty);

        }

        private void cboStatus_MouseClick(object sender, MouseEventArgs e)
        {
            error.SetError(cboStatus, string.Empty);

        }

        private void txtProfile_MouseClick(object sender, MouseEventArgs e)
        {
            error.SetError(txtProfile, string.Empty);

        }

        private void txtUsername_MouseClick(object sender, MouseEventArgs e)
        {
            error.SetError(txtUsername, string.Empty);

        }

        private void txtPassword_MouseClick(object sender, MouseEventArgs e)
        {
            error.SetError(txtPassword, string.Empty);

        }

        private void txtEmailAddress_MouseClick(object sender, MouseEventArgs e)
        {
            error.SetError(txtEmailAddress, string.Empty);

        }

        private void radMale_MouseClick(object sender, MouseEventArgs e)
        {
            error.SetError(label2, string.Empty);

        }

        private void radFemale_MouseClick(object sender, MouseEventArgs e)
        {
            error.SetError(label2, string.Empty);

        }

        private void chkVball_MouseClick(object sender, MouseEventArgs e)
        {
            error.SetError(label3, string.Empty);

        }

        private void chkBadminton_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void chkBadminton_MouseClick(object sender, MouseEventArgs e)
        {
            error.SetError(label3, string.Empty);

        }

        private void chkBball_MouseClick(object sender, MouseEventArgs e)
        {
            error.SetError(label3, string.Empty);

        }

        private void btnClear_Click(object sender, EventArgs e)
        {
            txtName.Clear();
            radFemale.Checked = false;
            radMale.Checked = false;

            chkBadminton.Checked = false;
            chkBball.Checked = false;
            chkVball.Checked = false;

            cboColor.SelectedIndex = -1;
            cboDegree.SelectedIndex = -1;

            txtSayings.Clear();
            txtUsername.Clear();
            txtPassword.Clear();
            txtEmailAddress.Clear();

            dtpBirthdate.Value = DateTime.Today;
            lblAge.Text = string.Empty;
            lblID.Text = string.Empty;

            
        }


        
    }
}
