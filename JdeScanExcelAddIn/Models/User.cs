using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.ComponentModel;
using JdeScanExcelAddIn.Models;
using JdeScanExcelAddIn.Static;
using System.Text.RegularExpressions;

namespace JdeScanExcelAddIn.Models
{
    public class User : Entity<User>
    {
        [DisplayName("ID")]
        public int UserId { get; set; }
        public override int Id
        {
            set => value = UserId;
            get => UserId;
        }
        [DisplayName("Imię")]
        [Browsable(false)]
        public string Name { get; set; }
        [DisplayName("Nazwisko")]
        [Browsable(false)]
        public string Surname { get; set; }
        [DisplayName("Imię i nazwisko")]
        public string FullName
        {
            get
            {
                return Name + " " + Surname;
            }
            set
            {
                string[] names = Regex.Split(value, " ");
                if (names.Count() > 2)
                {
                    Name = names[0];
                    Surname = string.Join(" ", names.Skip(1));
                }else if(names.Count() == 2)
                {
                    Name = names[0];
                    Surname = names[1];
                }
                else
                {
                    Name = names[0];
                }
            }
        }

        public string Password { get; set; }
        [DisplayName("Mechanik?")]
        public bool IsMechanic { get; set; }
        [DisplayName("Login MES")]
        public string MesLogin { get; set; }
        [DisplayName("Ostatnie logowanie")]
        public DateTime? LastLoggedOn { get; set; }

        public bool? IsArchived { get; set; }

    }
}
