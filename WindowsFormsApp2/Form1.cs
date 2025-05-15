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
using System.Xml.Linq;
using WindowsFormsApp2.Class;

namespace WindowsFormsApp2
{
    public partial class Form1 : Form
    {
       

        Form2 form2 = new Form2();
        Logs myLogs;
        public Form1()
        {
            InitializeComponent();
            
        }
        public Form1(Logs myLogs)
        {
            InitializeComponent();
            this.myLogs = myLogs;
            picProfile.Image = myLogs.Profile;
            picProfile.SizeMode = PictureBoxSizeMode.StretchImage;
            

        }
        private void Form1_Load(object sender, EventArgs e)
        {
            lblTime.Text = DateTime.Now.ToString("hh:mm:ss tt");
            lblDate.Text = DateTime.Now.ToString("MM/dd/yyyy");
            lblName.Text = myLogs.GlobalUser;

            
        }


        private void btnUpdate_Click(object sender, EventArgs e)
        {
            Workbook book = new Workbook();
            book.LoadFromFile(myLogs.FilePath);
            Worksheet sheet = book.Worksheets[0];

            

            bool errorRestriction = false;
            error.Clear();

            if (string.IsNullOrEmpty(txtName.Text))
            {
                error.SetError(txtName, "Name is required.");
                errorRestriction = true;
            }
            if (!radMale.Checked && !radFemale.Checked)
            {
                error.SetError(label2, "Select Gender.");
                errorRestriction = true;
            }
            if (!chkBadminton.Checked && !chkBball.Checked && !chkVball.Checked)
            {
                error.SetError(label3, "Select at least one hobby.");
                errorRestriction = true;
            }
            if (cboDegree.SelectedItem == null)
            {
                error.SetError(cboDegree, "Select a degree.");
                errorRestriction = true;
            }
            if (cboColor.SelectedItem == null)
            {
                error.SetError(cboColor, "Select a color.");
                errorRestriction = true;
            }
            if (string.IsNullOrEmpty(txtSayings.Text))
            {
                error.SetError(txtSayings, "Sayings is required.");
                errorRestriction = true;
            }
            if (dtpBirthdate.Value.Date > DateTime.Today)
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
            //if (cboStatus.SelectedItem == null)
            //{
            //    error.SetError(cboStatus, "Select a status.");
            //    errorRestriction = true;
            //}
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
            string sayings = txtSayings.Text;
            string username = txtUsername.Text;
            string password = txtPassword.Text;
            string email = txtEmailAddress.Text;
            string imagePath = txtProfile.Text;


            int age = myLogs.CalculateAge(dtpBirthdate.Value);
            lblAge.Text = age.ToString();

            bool userExists = false;
            int totalRows = sheet.Rows.Length;
            int currentRow = Convert.ToInt32(lblID.Text) + 2;

            for (int i = 2; i <= totalRows; i++)
            {
                if(currentRow == i) continue;
                string existsUser = sheet.Range[i, 8].Value;
                string existsPass = sheet.Range[i, 9].Value;

                if (existsUser == txtUsername.Text || existsPass == txtPassword.Text)
                {
                    userExists = true;
                    break;

                }

            }
            if (userExists)
            {
                MessageBox.Show("Username and Password already exist. Please provide another one.", "Duplication Found", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            sheet.Range[currentRow, 1].Value = name;
            sheet.Range[currentRow, 2].Value = gender;
            sheet.Range[currentRow, 3].Value = hobbies;
            sheet.Range[currentRow, 4].Value = degree;
            sheet.Range[currentRow, 5].Value = color;
            sheet.Range[currentRow, 6].Value = sayings;
            sheet.Range[currentRow, 7].Value = age.ToString();
            sheet.Range[currentRow, 8].Value = username;
            sheet.Range[currentRow, 9].Value = password;
            sheet.Range[currentRow, 11].Value = imagePath;
            sheet.Range[currentRow, 12].Value = email;
                
   
            book.SaveToFile(myLogs.FilePath);
       

            myLogs.insertLogs(myLogs.GlobalUser, "Updating Info");

            DataTable dt = sheet.ExportDataTable();
            form2.dgvData.DataSource = dt;
            MessageBox.Show("Successfully updated data.");


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

        private void txtName_TextChanged(object sender, EventArgs e)
        {

        }

        private void txtPassword_TextChanged(object sender, EventArgs e)
        {

        }

        private void guna2GradientPanel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void btnDashboard_Click(object sender, EventArgs e)
        {
            Dashboard dashboard = new Dashboard(myLogs);
            myLogs.insertLogs(myLogs.GlobalUser, "Visited Dashboard");
            dashboard.Show();
            this.Hide();
        }

        private void btnActiveStatus_Click(object sender, EventArgs e)
        {
            addStudent addStudent = new addStudent(myLogs);
            myLogs.insertLogs(myLogs.GlobalUser, "Visited Active Student");
            addStudent.Show();
            this.Hide();
        }

        private void btnInactiveStatus_Click(object sender, EventArgs e)
        {
            InActiveStudent InactiveStudent = new InActiveStudent(myLogs);
            myLogs.insertLogs(myLogs.GlobalUser, "Visited Inactive Student");
            InactiveStudent.Show();
            this.Hide();
        }

        private void btnAddStudent_Click(object sender, EventArgs e)
        {
            addStudent addStudent = new addStudent(myLogs);
            myLogs.insertLogs(myLogs.GlobalUser, "Visited Add Student");
            addStudent.Show();
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

        private void btnProfile_Click(object sender, EventArgs e)
        {
            Workbook book = new Workbook();
            book.LoadFromFile(myLogs.FilePath);
        
            Worksheet sheet = book.Worksheets[0];

            OpenFileDialog file = new OpenFileDialog();
            file.Filter = "Image Files|*.jpg;*.jpeg;*.png;*.bmp;*.gif";


            if (file.ShowDialog() == DialogResult.OK)
            {
                Image selectedImage = Image.FromFile(file.FileName);

                picProfile.Image = selectedImage;
                picProfile.SizeMode = PictureBoxSizeMode.StretchImage;

                txtProfile.Text = file.FileName;

                myLogs.Profile = selectedImage;
                myLogs.ImagePath = file.FileName;

                int row = sheet.Rows.Length + 1;

                sheet.Range[row, 11].Value = file.FileName;

                book.SaveToFile(myLogs.FilePath, ExcelVersion.Version2016);
           

            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void txtName_MouseClick(object sender, MouseEventArgs e)
        {
            error.SetError(txtName, string.Empty);
        }

        private void dtpBirthdate_MouseClick(object sender, MouseEventArgs e)
        {
            int age = myLogs.CalculateAge(dtpBirthdate.Value);
            lblAge.Text = age.ToString();
        }

        private void dtpBirthdate_ValueChanged(object sender, EventArgs e)
        {
            int age = myLogs.CalculateAge(dtpBirthdate.Value);
            lblAge.Text = age.ToString();
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

        private void chkBball_MouseClick(object sender, MouseEventArgs e)
        {
            error.SetError(label3, string.Empty);

        }

        private void chkBadminton_MouseClick(object sender, MouseEventArgs e)
        {
            error.SetError(label3, string.Empty);

        }

        private void cboDegree_MouseClick(object sender, MouseEventArgs e)
        {
            error.SetError(cboDegree, string.Empty);

        }

        private void cboColor_MouseClick(object sender, MouseEventArgs e)
        {
            error.SetError(cboColor, string.Empty);

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

        private void cboStatus_MouseClick(object sender, MouseEventArgs e)
        {
            //error.SetError(cboStatus, string.Empty);

        }

        private void btnProfile_Click_1(object sender, EventArgs e)
        {
            OpenFileDialog file = new OpenFileDialog();
            file.Filter = "Image Files|*.jpg;*.jpeg;*.png;*.bmp;*.gif";


            if (file.ShowDialog() == DialogResult.OK)
            {
                txtProfile.Text = file.FileName;


            }
        }
    }
}
