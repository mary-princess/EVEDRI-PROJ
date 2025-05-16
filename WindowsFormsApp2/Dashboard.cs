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
    public partial class Dashboard : Form
    {
        public Dashboard()
        {
            InitializeComponent();
        }
        Logs myLogs;
        public Dashboard(Logs myLogs)
        {
            InitializeComponent();
            this.myLogs = myLogs;
            picProfile.Image = myLogs.Profile;
            picProfile.SizeMode = PictureBoxSizeMode.StretchImage;
           
        }


        private void Dashboard_Load(object sender, EventArgs e)
        {
            lblTime.Text = DateTime.Now.ToString("hh:mm:ss tt");
            lblDate.Text = DateTime.Now.ToString("MM/dd/yyyy");
            lblName.Text = myLogs.GlobalUser;

            lblActiveStudent.Text = showCount(10, "1").ToString();
            lblBadminton.Text = showCount(3, "Badminton").ToString();
            lblBasketball.Text = showCount(3, "Basketball").ToString();
            lblBlue.Text = showCount(5, "Blue").ToString();
            lblBscpe.Text = showCount(4, "BSCpE").ToString();
            lblBscs.Text = showCount(4, "BSCS").ToString();
            lblBsit.Text = showCount(4, "BSIT").ToString();
            lblFemaleStudent.Text = showCount(2, "Female").ToString();
            lblInactiveStudent.Text = showCount(10, "0").ToString();
            lblMaleStudent.Text = showCount(2, "Male").ToString();
            lblRed.Text = showCount(5, "Red").ToString();
            lblVolleyball.Text = showCount(3, "Volleyball").ToString();
            lblYellow.Text = showCount(5, "Yellow").ToString();
           

        }
        public int showCount(int c, string value)
        {
            Workbook wb = new Workbook();
            wb.LoadFromFile(myLogs.FilePath);
            Worksheet sh = wb.Worksheets[0];
            int row = sh.Rows.Length;
            int counter = 1;

            for (int i = 2; i < row; i++)
            {
                if (sh.Range[i, c].Value == value)
                {
                    counter++;
                }
            }
            return counter;
        }
        
        private void btnActiveStatus_Click(object sender, EventArgs e)
        {
           Form2 activestudent = new Form2(myLogs);
            myLogs.insertLogs(myLogs.GlobalUser, "Visited Active Student");

            activestudent.Show();
            this.Hide();
        }

        private void btnDashboard_Click(object sender, EventArgs e)
        {

        }

        private void btnInactiveStatus_Click(object sender, EventArgs e)
        {
            InActiveStudent inActiveStudent = new InActiveStudent(myLogs);
            myLogs.insertLogs(myLogs.GlobalUser, "Visited Inactive Student");
            inActiveStudent.Show();
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
            frmLogs log = new frmLogs(myLogs);
            myLogs.insertLogs(myLogs.GlobalUser, "Visited Logs");
            log.Show();
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

        private void guna2GradientPanel5_Paint(object sender, PaintEventArgs e)
        {

        }

        private void pictureBox8_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("Are you sure you want to delete this data?", "Confirm Delete", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (result == DialogResult.No)
            {
                return;
            }
            Application.Exit();
        }
    }
}
