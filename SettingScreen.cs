using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;

using System.Windows.Forms;

namespace TableLayoutPanelSample
{
    public partial class SettingScreen : Form
    {
        OleDbDataAdapter adapter;
        DataSet ds;
        int serial_nos = 0, Df_number, DDf_number;
        DataTable dt;
        int disp_product, disp_vendor, _disp_serial, disp_date, bar_product, bar_date, bar_serial, bar_vendor,dis_descrp;
        String date_format;
        private OleDbConnection connection = new OleDbConnection();
        public SettingScreen()
        {
            InitializeComponent();
            connection.ConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Application.StartupPath + "/dataReader.accdb;Persist Security Info=True";
            setData();
            TopMost = true;
        }

        private void setData()
        {
            String query = "select * from print_setting";
            adapter = new OleDbDataAdapter(query, connection);
            dt = new DataTable();//student-> table name in stud.accdb file
            adapter.Fill(dt);
            foreach (DataRow dataRow in dt.Rows)
            {

                checkBox4.Checked = getValueOf(dataRow.Field<Int32>("display_product"));
                checkBox2.Checked = getValueOf(dataRow.Field<Int32>("display_date"));
                checkBox3.Checked = getValueOf(dataRow.Field<Int32>("display_vendor"));

                checkBox1.Checked = getValueOf(dataRow.Field<Int32>("display_serial"));
                checkBox10.Checked = getValueOf(dataRow.Field<Int32>("display_vaipl"));

                checkBox5.Checked = getValueOf(dataRow.Field<Int32>("bar_product"));

                checkBox7.Checked = getValueOf(dataRow.Field<Int32>("bar_date"));

                checkBox6.Checked = getValueOf(dataRow.Field<Int32>("bar_vendor"));

                checkBox8.Checked = getValueOf(dataRow.Field<Int32>("bar_serial"));
                checkBox9.Checked = getValueOf(dataRow.Field<Int32>("display_desc"));

                dis_shift.Checked = getValueOf(dataRow.Field<Int32>("display_rev_level"));
                checkBox12.Checked = getValueOf(dataRow.Field<Int32>("bar_rev_level"));

                switch (dataRow.Field<Int32>("date_format"))
                {
                    case 1:
                        radioButton1.Checked = true;
                        Df_number = 1;
                        break;
                    case 2:
                        radioButton2.Checked = true;
                        Df_number = 2;
                        break;
                    case 3:
                        radioButton3.Checked = true;
                        Df_number = 3;
                        break;

                }
                switch (dataRow.Field<Int32>("dis_date_frmt"))
                {
                    case 1:
                        radioButton6.Checked = true;
                        DDf_number = 1;
                        break;
                    case 2:
                        radioButton5.Checked = true;
                        DDf_number = 2;
                        break;
                    case 3:
                        radioButton4.Checked = true;
                        DDf_number = 3;
                        break;

                }
            }
        }

        private void tableLayoutPanel1_Paint(object sender, PaintEventArgs e)
        {

        }
        public int getValueOfBool(bool val)
        {
            if (val)
                return 1;
            else
                return 0;
        }
        public bool getValueOf(int val)
        {
            if (val == 1)
                return true;
            else
                return false;
        }

        private void checkBox8_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void xButton1_Click(object sender, EventArgs e)
        {
            connection.Open();
            OleDbCommand command = new OleDbCommand();
            command.Connection = connection;
            String query = "Update print_setting set `display_product`='" + getValueOfBool(checkBox4.Checked) + "',`display_date`='" + getValueOfBool(checkBox2.Checked) +
                "',`display_serial`='" + getValueOfBool(checkBox1.Checked) + "',`display_vendor`='" + getValueOfBool(checkBox3.Checked) +
                "',`bar_product`='" + getValueOfBool(checkBox5.Checked) + "',`bar_date`='" + getValueOfBool(checkBox7.Checked) + "',`bar_vendor`='" + getValueOfBool(checkBox6.Checked) +
                "',`bar_serial`='" + getValueOfBool(checkBox8.Checked) + "',`date_format`='" + Df_number + "',`dis_date_frmt`='" + DDf_number +
                "',`display_desc`='" + getValueOfBool(checkBox9.Checked) + "',`display_vaipl`='" + getValueOfBool(checkBox10.Checked) +
                "',`display_rev_level`='" + getValueOfBool(dis_shift.Checked) + "',`bar_rev_level`='" + getValueOfBool(checkBox12.Checked) + "'";
            command.CommandText = query;
            command.ExecuteNonQuery();
            connection.Close();
            MessageBox.Show("Update Succcessful");
            this.Hide();

        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton1.Checked)
            {
                Df_number = 1;
            }
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton2.Checked)
            {
                Df_number = 2;
            }
        }

        private void radioButton3_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton3.Checked)
            {
                Df_number = 3;
            }
        }

        private void radioButton6_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton6.Checked)
            {
                DDf_number = 1;
            }
        }

        private void radioButton5_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton5.Checked)
            {
                DDf_number = 2;
            }
        }

        private void radioButton4_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton4.Checked)
            {
                DDf_number = 3;
            }
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {

            groupBox4.Enabled = checkBox2.Checked;

        }

        private void checkBox7_CheckedChanged(object sender, EventArgs e)
        {
            groupBox3.Enabled = checkBox7.Checked;
        }

        private void xButton2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void xButton4_Click(object sender, EventArgs e)
        {

        }
    }
}
