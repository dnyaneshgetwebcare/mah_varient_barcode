using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace TableLayoutPanelSample
{
    public partial class PasswordForm : Form
    {
        Form requestForm;
        public PasswordForm()
        {
            InitializeComponent();
        }
        public PasswordForm(Form request_form)
        {
            InitializeComponent();
            this.requestForm = request_form;
        }
        private void xButton1_Click(object sender, EventArgs e)
        {
            if (textBox1.Text.ToString().Equals("Password#2025"))
            {

                requestForm.Show();
                this.Hide();
            }
            else
            {
                MessageBox.Show("Invalid Password");
            }
        }

        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                xButton1.PerformClick();
            }
        }
    }
}