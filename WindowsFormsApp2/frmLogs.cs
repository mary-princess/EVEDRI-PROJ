using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics.Eventing.Reader;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsFormsApp2.Class
{
    public partial class frmLogs : Form
    {
        public frmLogs()
        {
            InitializeComponent();
        }
        Logs myLogs;

        public frmLogs(Logs logs)
        {
            InitializeComponent();
            this.myLogs = logs;
            picProfile.Image = myLogs.Profile;
            picProfile.SizeMode = PictureBoxSizeMode.StretchImage;
        }

        private void frmLogs_Load(object sender, EventArgs e)
        {
            lblTime.Text = DateTime.Now.ToString("hh:mm:ss tt");
            lblDate.Text = DateTime.Now.ToString("MM/dd/yyyy");
            lblName.Text = myLogs.GlobalUser;
            myLogs.showLogs(dgvLogs);

           
        }

        private void dgvLogs_CellContentClick(object sender, DataGridViewCellEventArgs e)
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
            Form2 activeStudent = new Form2(myLogs);
            myLogs.insertLogs(myLogs.GlobalUser, "Visited Active Student");
            activeStudent.Show();
            this.Hide();
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

        }

        private void btnLogout_Click(object sender, EventArgs e)
        {
            Login login = new Login(myLogs);
            myLogs.insertLogs(myLogs.GlobalUser, "Log Out");
            login.Show();
            this.Hide();
        }

        private void btnSearch_Click(object sender, EventArgs e)
        {

            string searchValue = txtSearch.Text.Trim().ToLower();
            dgvLogs.ClearSelection();

            if (string.IsNullOrEmpty(searchValue))
            {
                MessageBox.Show("Please enter a value to search.", "Search", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            int columnIndex = -1;
            if (cboSearch.SelectedIndex == 0)
            {
                columnIndex = 0; // User
            }
            else if (cboSearch.SelectedIndex == 1)
            {
                columnIndex = 1; // Message
            }
            
            if(columnIndex == -1)
            {
                MessageBox.Show("Please select a search filter.", "Search", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            bool matchFound = false;
            foreach (DataGridViewRow row in dgvLogs.Rows)
            {
                if(row.IsNewRow) continue; // Skip the new row placeholder

                if (row.Cells[columnIndex].Value != null && row.Cells[columnIndex].Value.ToString().ToLower().Contains(searchValue))
                {
                    row.Selected = true;
                    if (!matchFound)
                    {
                        dgvLogs.FirstDisplayedScrollingRowIndex = row.Index; // Scroll to the first match
                    }
                    matchFound = true;
                }
            }
            if(!matchFound)
            {
                MessageBox.Show("No matching records found.", "Search", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void cboSearch_SelectedIndexChanged(object sender, EventArgs e)
        {
            dgvLogs.ClearSelection();
        }

        private void cboSearch_MouseClick(object sender, MouseEventArgs e)
        {
            dgvLogs.ClearSelection();
        }
    }
}
