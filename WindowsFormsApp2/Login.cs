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
    public partial class Login : Form
    {
        public Login()
        {
            InitializeComponent();
        }
        Logs logs = new Logs();  
        public Login(Logs logs)
        {
            InitializeComponent();
            this.logs = logs;
        }

        private void Login_Load(object sender, EventArgs e)
        {

        }

        

        private void guna2Button1_Click(object sender, EventArgs e)
        {
            Workbook book = new Workbook();
            book.LoadFromFile(logs.FilePath);


            Worksheet sheet = book.Worksheets[0];

            bool log = false;
            int row = sheet.Rows.Length;
            for (int i = 2; i <= row; i++)
            {
                if (sheet.Range[i, 8].Value == txtUsername.Text &&
                    sheet.Range[i, 9].Value == txtPassword.Text)
                {
                    logs.Profile = Image.FromFile(sheet.Range[i, 11].Value);

                    logs.GlobalUser = sheet.Range[i, 8].Value;
                    logs.insertLogs(logs.GlobalUser, "Log In");

                    
                        
                    logs.Profile = Image.FromFile(sheet.Range[i, 11].Value);
                    

                    log = true;
                    break;

                }
                else
                {
                    log = false;
                }

            }

            if (log == true)
            {
                MessageBox.Show("Successfully Login", "Login", MessageBoxButtons.OK, MessageBoxIcon.Information);
                Dashboard dashboard = new Dashboard(logs);
                dashboard.Show();
                this.Hide();
            }
            else
            {
                MessageBox.Show("Incorrect username or password", "Login Failed", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
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
