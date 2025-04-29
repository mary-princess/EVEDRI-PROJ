using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WindowsFormsApp2.Class
{
    public class Usercs
    {
        public string GlobalUser {  get; set; }

        public Image Profile {  get; set; }
        public string ImagePath { get; set; }

        public string FilePath { get; set; }    
        public Usercs() { }

        public Usercs(Usercs usercs)
        {
            FilePath = @"C:\\Users\\ACT-STUDENT\\Desktop\Book1.xlsx"; //school
            //FilePath = @"C:\Users\User\OneDrive\Desktop\Book1.xlsx"; //laptop
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
            return emailAddress.EndsWith("@gmail.com", StringComparison.OrdinalIgnoreCase);
        }
    }
}
