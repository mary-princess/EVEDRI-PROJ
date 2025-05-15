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

            Form2 form = (Form2)Application.OpenForms["Form2"];
            bool userExists = false;
            int totalRows = sheet.Rows.Length;

            for (int i = 2; i <= totalRows; i++)
            {
                string existsUser = sheet.Range[i, 8].Value;
                string existsPass = sheet.Range[i, 9].Value;

                if (existsUser == txtUsername.Text || existsPass == txtPassword.Text)
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

            int r = form2.dgvData.CurrentCell.RowIndex;

            if (!string.IsNullOrEmpty(txtName.Text)) form2.dgvData.Rows[r].Cells[0].Value = txtName.Text;
            

            string gender = "";
            if (radFemale.Checked == true)
            {
                form2.dgvData.Rows[r].Cells[1].Value = radFemale.Text;
                gender = radFemale.Text;
            }
            else if (radMale.Checked == true)
            {
                form2.dgvData.Rows[r].Cells[1].Value = radMale.Text;
                gender = radMale.Text;
            }


            string hobbies = "";
            if (chkVball.Checked == true) hobbies += chkVball.Text + " , ";
            if (chkBball.Checked == true) hobbies += chkBball.Text + " , ";
            if (chkBadminton.Checked == true) hobbies += chkBadminton.Text + " , ";
            form2.dgvData.Rows[r].Cells[2].Value = hobbies.Trim();

            if (cboColor.SelectedItem != null) form2.dgvData.Rows[r].Cells[3].Value = cboColor.Text;
            if (cboDegree.SelectedItem != null) form2.dgvData.Rows[r].Cells[4].Value = cboDegree.Text;
            if (!string.IsNullOrEmpty(txtSayings.Text)) form2.dgvData.Rows[r].Cells[5].Value = txtSayings.Text;

            DateTime birthdate = dtpBirthdate.Value;
            form2.dgvData.Rows[r].Cells[6].Value = birthdate.ToString();

            if (!string.IsNullOrEmpty(txtUsername.Text)) form2.dgvData.Rows[r].Cells[7].Value = txtUsername.Text;
            if (!string.IsNullOrEmpty(txtPassword.Text)) form2.dgvData.Rows[r].Cells[8].Value = txtPassword.Text;
            if (cboStatus.SelectedItem != null) form2.dgvData.Rows[r].Cells[9].Value = cboStatus.Text;
            string imagePath = txtProfile.Text;
            if (!string.IsNullOrEmpty(imagePath)) form2.dgvData.Rows[r].Cells[10].Value = imagePath;

            int age = myLogs.CalculateAge(birthdate);
            lblAge.Text = age.ToString();

            int row = (Convert.ToInt32(lblID.Text)) + 2;

            sheet.Range[row, 1].Value = txtName.Text;
            sheet.Range[row, 2].Value = gender;
            sheet.Range[row, 3].Value = hobbies;
            sheet.Range[row, 4].Value = cboDegree.Text;
            sheet.Range[row, 5].Value = cboColor.Text;
            sheet.Range[row, 6].Value = txtSayings.Text;
            sheet.Range[row, 7].Value = myLogs.Age.ToString();
            sheet.Range[row, 8].Value = txtUsername.Text;
            sheet.Range[row, 9].Value = txtPassword.Text;
            sheet.Range[row, 10].Value = cboStatus.Text;
            sheet.Range[row, 11].Value = imagePath;
            sheet.Range[row, 12].Value = cboStatus.Text;


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
            error.SetError(cboStatus, string.Empty);

        }
    }
}
