using JdeScanExcelAddIn.Models;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace JdeScanExcelAddIn
{
    public partial class frmLogin : Form
    {
        UsersKeeper Keeper = new UsersKeeper();

        public frmLogin(UsersKeeper keeper)
        {
            InitializeComponent();
            Keeper = keeper;
        }

        private void formLoaded(object sender, EventArgs e)
        {
            cmbUsers.DataSource = Keeper.Items;
            cmbUsers.DisplayMember = "FullName";
            cmbUsers.ValueMember = "UserId";
            txtPassword.Text = "Hasło";
            txtPassword.ForeColor = Color.Gray;
        }

        private void txtPassword_Enter(object sender, EventArgs e)
        {
            if (txtPassword.Text == "Hasło")
            {
                txtPassword.Text = "";
                txtPassword.ForeColor = Color.Black;
                txtPassword.PasswordChar = '*';
            }
        }

        private void txtPassword_Leave(object sender, EventArgs e)
        {
            if (txtPassword.Text == "")
            {
                txtPassword.Text = "Hasło";
                txtPassword.ForeColor = Color.Gray;
                txtPassword.PasswordChar = '\0';
            }
        }

        private void Login(object sender, EventArgs e)
        {

            Login();
        }

        private void Login()
        {
            if (txtPassword.ForeColor == Color.Black && cmbUsers.SelectedItem != null)
            {
                int UserId = Convert.ToInt32(cmbUsers.SelectedValue.ToString());
                if (Keeper.Items.Where(u => u.UserId == UserId && u.Password == txtPassword.Text).Any())
                {
                    //login and password OK
                    Settings.CurrentUser = Keeper.Items.Where(u => u.UserId == UserId && u.Password == txtPassword.Text).FirstOrDefault();
                    User _this = Keeper.Items.Where(u => u.UserId == UserId).FirstOrDefault();
                    this.DialogResult = DialogResult.OK;
                    this.Close();
                }
                else
                {
                    MessageBox.Show("Podano błędne hasło!");
                }

            }
            else
            {
                MessageBox.Show("Nie wybrano użytkownika lub nie podano hasła! Najpierw wybierz użytkownika i podaj hasło.");
            }
        }

        private void frmLogin_FormClosing(object sender, FormClosingEventArgs e)
        {
        }

        private void tlpMain_Paint(object sender, PaintEventArgs e)
        {

        }

        private void txtPassword_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                Login();
            }
        }
    }
}
