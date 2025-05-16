using Guna.UI2.WinForms.Enums;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace WindowsFormsApp2.Class
{
    public class Usercs
    {
        public string GlobalUser {  get; set; }

        public Image Profile {  get; set; }
        public string ImagePath { get; set; }

        public string FilePath { get; set; }    
        public Usercs() 
        {
            //FilePath = @"C:\Users\User\OneDrive\Desktop\Book1.xlsx"; //laptop
            FilePath = @"C:\Users\ACT-STUDENT\Desktop\Book1.xlsx"; //school
        }

        public Usercs(Usercs usercs)
        {
            this.GlobalUser = usercs.GlobalUser;
            this.FilePath = usercs.FilePath;
            this.ImagePath  = usercs.ImagePath;
            this.Profile = usercs.Profile;
            
        }

        public int Age { get; set; }

        public int CalculateAge(DateTime birthdate)
        {
            var today = DateTime.Today;
            var age = today.Year - birthdate.Year;
            if (birthdate > today.AddYears(-age))
                age--;
            return age;
        }
        public bool ValidateEmailAddress(string emailAddress)
        {
            if (string.IsNullOrEmpty(emailAddress))
            {
                return false;
            }

            // More comprehensive regular expression for email validation
            string emailRegex = @"^[a-zA-Z0-9.!#$%&'*+/=?^_`{|}~-]+@[a-zA-Z0-9](?:[a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?(?:\.[a-zA-Z0-9](?:[a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?)*$";

            return Regex.IsMatch(emailAddress, emailRegex);
        }
    }
}
