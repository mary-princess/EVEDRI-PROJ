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
            //if (!string.IsNullOrEmpty(myLogs.ImagePath) && File.Exists(myLogs.ImagePath))
            //{
            //    picProfile.Image = Image.FromFile(myLogs.ImagePath);
            //    picProfile.SizeMode = PictureBoxSizeMode.StretchImage;
            //}

        }
        private void Form1_Load(object sender, EventArgs e)
        {
            lblTime.Text = DateTime.Now.ToString("hh:mm:ss tt");
            lblDate.Text = DateTime.Now.ToString("MM/dd/yyyy");
            lblName.Text = myLogs.GlobalUser;

            
        }


        private void btnInsert_Click(object sender, EventArgs e)
        {
            //Workbook book = new Workbook();
            //book.LoadFromFile(@"C:\Users\User\OneDrive\Desktop\Book1.xlsx");
            //Worksheet sheet = book.Worksheets[0];

            //string name = "",
            //    gender = "",
            //    hobbies = "",
            //    favoriteColor = "",
            //    sayings = "",
            //    username = "",
            //    password = "";
            //if (!string.IsNullOrEmpty(txtName.Text))
            //{
            //    name += txtName.Text;
            //}
            //if (radFemale.Checked == true)
            //{
            //    gender += radFemale.Text;
            //}
            //if (radMale.Checked == true)
            //{
            //    gender += radMale.Text;
            //}

            //if (chkVball.Checked == true)
            //{
            //    hobbies += chkVball.Text + " , ";
            //}
            //if (chkBball.Checked == true)
            //{
            //    hobbies += chkBball.Text + " , ";
            //}
            //if (chkBadminton.Checked == true)
            //{
            //    hobbies += chkBadminton.Text + " , ";
            //}

            //if (cboColor.SelectedItem != null)
            //{
            //    favoriteColor += cboColor.Text;
            //}

            //if (!string.IsNullOrEmpty(txtSayings.Text))
            //{
            //    sayings += txtSayings.Text;

            //}
            //if (!string.IsNullOrEmpty(txtUsername.Text))
            //{
            //    username += txtUsername.Text;
            //}
            //if (!string.IsNullOrEmpty(txtPassword.Text))
            //{
            //    password += txtPassword.Text;
            //}


            ////form2.insertTable(++id, name, gender, hobbies, favoriteColor, sayings, username, password);


            //int row = sheet.Rows.Length + 1;

            //sheet.Range[row, 1].Value = name;
            //sheet.Range[row, 2].Value = gender;
            //sheet.Range[row, 3].Value = hobbies;
            //sheet.Range[row, 4].Value = favoriteColor;
            //sheet.Range[row, 5].Value = sayings;
            //sheet.Range[row, 6].Value = username;
            //sheet.Range[row, 7].Value = password;

            //book.SaveToFile(@"C:\Users\User\OneDrive\Desktop\Book1.xlsx", ExcelVersion.Version2016);
            //myLogs.insertLogs(username, "Inserting Info");
            //DataTable dt = sheet.ExportDataTable();
            //form2.dgvData.DataSource = dt;
            //MessageBox.Show("Successfully added data.");


        }

        private void btnUpdate_Click(object sender, EventArgs e)
        {
            Workbook book = new Workbook();
            book.LoadFromFile(myLogs.FilePath);
            Worksheet sheet = book.Worksheets[0];

            Form2 form = (Form2)Application.OpenForms["Form2"];
           

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

            dtpBirthdate.Value = DateTime.Today;
            lblAge.Text = string.Empty;
            lblID.Text = string.Empty;

            picProfile.Image = null;
            txtProfile.Clear();
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
            Login login = new Login(myLogs);
            myLogs.insertLogs(myLogs.GlobalUser, "Log Out");
            login.Show();
            this.Hide();
        }

        private void btnProfile_Click(object sender, EventArgs e)
        {
            Workbook book = new Workbook();
            book.LoadFromFile(myLogs.FilePath);
            //book.LoadFromFile(@"C:\Users\User\OneDrive\Desktop\Book1.xlsx");
            //book.LoadFromFile(@"C:\Users\ACT-STUDENT\Desktop\Book1.xlsx");

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
                //book.SaveToFile(@"C:\Users\User\OneDrive\Desktop\Book1.xlsx", ExcelVersion.Version2016);
                //book.SaveToFile(@"C:\Users\ACT-STUDENT\Desktop\Book1.xlsx", ExcelVersion.Version2016);

            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
