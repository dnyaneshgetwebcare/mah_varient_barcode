using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Drawing.Printing;
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
            connection.ConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Application.StartupPath + "/voss_3_dataReader.accdb;Persist Security Info=True";
            setPrinter();
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
                disp_serial.Checked = getValueOf(dataRow.Field<Int32>("display_serial"));
                display_date.Checked = getValueOf(dataRow.Field<Int32>("display_date"));
                disp_prod_code.Checked = getValueOf(dataRow.Field<Int32>("display_product"));
                disp_vend_code.Checked = getValueOf(dataRow.Field<Int32>("display_vendor"));
                disp_desc.Checked = getValueOf(dataRow.Field<Int32>("display_desc"));
                disp_part_id.Checked = getValueOf(dataRow.Field<Int32>("disp_part_id"));
                dis_shift.Checked = getValueOf(dataRow.Field<Int32>("disp_shift"));
                disp_cust_rev.Checked = getValueOf(dataRow.Field<Int32>("display_rev_level"));


                bar_serial_no.Checked = getValueOf(dataRow.Field<Int32>("bar_serial"));
                barcode_date.Checked = getValueOf(dataRow.Field<Int32>("bar_date"));
                bar_productcode.Checked = getValueOf(dataRow.Field<Int32>("bar_product"));
                bar_vend_code.Checked = getValueOf(dataRow.Field<Int32>("bar_vendor"));
                barcode_shift.Checked = getValueOf(dataRow.Field<Int32>("bar_shift"));
                bar_cust_rev.Checked = getValueOf(dataRow.Field<Int32>("bar_rev_level"));
                bar_part_id.Checked = getValueOf(dataRow.Field<Int32>("bar_part_id"));
                cb_print_list.Text = dataRow.Field<String>("printer_name");

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


            String shift_query = "select * from shift_details";
            OleDbDataAdapter shif_adapter = new OleDbDataAdapter(shift_query, connection);
            dt = new DataTable();//student-> table name in stud.accdb file
            shif_adapter.Fill(dt);
            foreach (DataRow dataRow in dt.Rows)
            {
                String shift_name = dataRow.Field<String>("shift_name");
                String shift_start_time = dataRow.Field<String>("start_time");
                String shift_end_time = dataRow.Field<String>("end_time");
                switch (dataRow.Field<Int32>("ID"))
                {
                    case 2:
                        tb_name_shift1.Text = shift_name;
                        tp_start_shift1.Value = DateTime.Parse(shift_start_time);
                        tp_end_shift.Value = DateTime.Parse(shift_end_time);
                        break;
                    case 3:
                        tb_name_shift2.Text = shift_name;
                        tp_start_shift2.Value = DateTime.Parse(shift_start_time);
                        tp_end_shift2.Value = DateTime.Parse(shift_end_time);
                        break;
                    case 4:
                        tb_name_shift3.Text = shift_name;
                        tp_start_shift3.Value = DateTime.Parse(shift_start_time);
                        tp_end_shift3.Value = DateTime.Parse(shift_end_time);
                        break;
                }
            }

            String number_query = "select * from serial_nos";
            OleDbDataAdapter number_adapter = new OleDbDataAdapter(number_query, connection);
            dt = new DataTable();//student-> table name in stud.accdb file
            number_adapter.Fill(dt);
            foreach (DataRow dataRow in dt.Rows)
            {
                int cur_serial_no = dataRow.Field<Int32>("current_serial_nos");
                tb_current_no.Text =cur_serial_no.ToString();

            }
        }
        private void setPrinter()
        {
            PrinterSettings.StringCollection printerNames = PrinterSettings.InstalledPrinters;
            // Alternatively, you can add the printer names to a list
            List<string> printerList = new List<string>();
            foreach (string printerName in printerNames)
            {
              //  printerList.Add(printerName);
                cb_print_list.Items.Add(printerName);
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
            String query = "Update print_setting set `display_product`='" + getValueOfBool(disp_prod_code.Checked) 
                + "',`display_date`='" + getValueOfBool(display_date.Checked) 
                + "',`display_serial`='" + getValueOfBool(disp_serial.Checked) 
                + "',`display_vendor`='" + getValueOfBool(disp_vend_code.Checked) 
                + "',`bar_product`='" + getValueOfBool(bar_productcode.Checked) 
                + "',`bar_date`='" + getValueOfBool(barcode_date.Checked) 
                + "',`bar_vendor`='" + getValueOfBool(bar_vend_code.Checked) 
                + "',`bar_serial`='" + getValueOfBool(bar_serial_no.Checked) 
                + "',`date_format`='" + Df_number 
                + "',`dis_date_frmt`='" + DDf_number 
                +"',`display_desc`='" + getValueOfBool(disp_desc.Checked) 
                + "',`disp_part_id`='" + getValueOfBool(disp_part_id.Checked) 
                + "',`display_rev_level`='" + getValueOfBool(disp_cust_rev.Checked) 
                + "',`bar_rev_level`='" + getValueOfBool(bar_cust_rev.Checked) 
                + "',`bar_part_id`='" + getValueOfBool(bar_part_id.Checked) 
                + "',`printer_name`='" + cb_print_list.Text.ToString() 
                + "',`bar_shift`='" + getValueOfBool(barcode_shift.Checked) 
                + "',`disp_shift`='" + getValueOfBool(dis_shift.Checked) + "';";
            String query1 = "Update shift_details set `shift_name` ='" + tb_name_shift1.Text.ToString() + "' ,`start_time` = '" + tp_start_shift1.Value.ToString("HH:mm") + "', `end_time` = '" + tp_end_shift.Value.ToString("HH:mm") + "' where ID = 2;";
            String query2 = "Update shift_details set `shift_name` ='"+tb_name_shift2.Text.ToString()+"' ,`start_time` = '"+tp_start_shift2.Value.ToString("HH:mm") +"', `end_time` = '"+ tp_end_shift2.Value.ToString("HH:mm") +"' where ID = 3;";
           
            String query3 = "Update shift_details set `shift_name` ='" + tb_name_shift3.Text.ToString() + "' ,`start_time` = '" + tp_start_shift3.Value.ToString("HH:mm") + "', `end_time` = '" + tp_end_shift3.Value.ToString("HH:mm") + "' where ID = 4;";
            String serial_query = "Update serial_nos set `current_serial_nos` ='" + tb_current_no.Text.ToString() + "';";
            command.CommandText = query;
            command.ExecuteNonQuery();
            command.CommandText =  query1;
            command.ExecuteNonQuery();
            command.CommandText = query2;
            command.ExecuteNonQuery();
            command.CommandText = query3;
            command.ExecuteNonQuery();
            connection.Close();
            MessageBox.Show("Update Succcessful");
            this.Hide();

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void label4_Click(object sender, EventArgs e)
        {

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

            groupBox4.Enabled = display_date.Checked;

        }

        private void checkBox7_CheckedChanged(object sender, EventArgs e)
        {
            groupBox3.Enabled = barcode_date.Checked;
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
