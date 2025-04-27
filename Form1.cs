using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using System.IO.Ports;
using System.Threading;
using System.Text.RegularExpressions;


namespace TableLayoutPanelSample
{
    public partial class Form1 : Form
    {
        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);
        }
        private OleDbConnection connection = new OleDbConnection();
        public Form1()
        {
            InitializeComponent();
           // connection.ConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\c13\Documents\Visual Studio 2008\Projects\DataReader1\DataReader1\Resources\dataReader.accdb;Persist Security Info=True";
            connection.ConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\c13\Documents\Visual Studio 2008\Projects\TableLayoutPanelSample\dataReader.accdb;Persist Security Info=True";
           /* string[] ports = SerialPort.GetPortNames();
            foreach (string port in ports)
            {
                comboBox1.Items.Add(port);
            }
            * */
            // FormBorderStyle = FormBorderStyle.None;
            WindowState = FormWindowState.Maximized;
           // TopMost = true;
        }

      

       
        private void quitMenu_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Do you really wish to exit", "Quit Verification", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                return;
            Form1 f4 = new Form1();
            Form2 f2 = new Form2();
            Form2 f3 = new Form2();
            f4.Close();
            f2.Close();
            f3.Close();
            this.Close();
        }

        private void portSettingMenu_Click(object sender, EventArgs e)
        {
            Form2 f4 = new Form2();
            this.Hide();
            f4.Show();
        }

        private void exportDataMenu_Click(object sender, EventArgs e)
        {
            Form2 f3 = new Form2();
            this.Hide();
            f3.Show();
        }

        private void tableLayoutPanel2_Paint(object sender, PaintEventArgs e)
        {

        }

        private void comportMenu1_Click(object sender, EventArgs e)
        {
            Form2 f2 = new Form2();
            f2.Show();

        }
        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            Environment.Exit(0);
        }

        private void pageSetupToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form2 frm5 = new Form2();
            frm5.Show();
        }
    }
}
