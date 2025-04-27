using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Drawing.Printing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace TableLayoutPanelSample
{
    public partial class Form2 : Form
    {
        OleDbDataAdapter adapter;
        bool display_product = false, display_serial = false, display_date = false,
            display_vendor = false, bar_product = false, bar_vendor = false, 
            bar_date = false, bar_serial = false, display_descripton = false,
            bar_rev_level = false,  display_rev_level = false, display_vailp_rev=false;
        DataSet ds;
        String datefmt, display_date_fmt;
        bool preview_bool = true;
        int serial_nos = 0;
        DataTable dt;
        List<String> product_code_list, vendor_code;
        private OleDbConnection connection = new OleDbConnection();
        public Form2()
        {
            InitializeComponent();
            connection.ConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Application.StartupPath + "/voss_3_dataReader.accdb;Persist Security Info=True";
            //deleteOldRecords();
            setAutocomplete();
            loadData();

        }
        public void loadData()
        {
            String query = "select id as `Sr no`,date_entry as `Date`,time_entry as `Time`,barcode as `Barcode`,product_code as `Product Code`,serial_number as `Serial Number`,vendor_code as `Vendor Code`,descrip as `Description` from report where `date_entry`=#" + DateTime.Now.ToString("dd/MM/yyyy") + "# order by id desc";
            adapter = new OleDbDataAdapter(query, connection);
            dt = new DataTable();//student-> table name in stud.accdb file
            adapter.Fill(dt);
            DateTime theDate;
            theDate = DateTime.Now;
          // String temp= theDate.ToString("ddMMyy hh:mm tt");
            dataGridView1.DataSource = dt;
            //dataGridView1.Columns.Add("srn_no", "Sr No");
            //dataGridView1.Columns.Add("date_entry", "Date");
            //dataGridView1.Columns.Add("barcode", "Barcode");
            //dataGridView1.Columns.Add("product_code", "Product Code");
            //dataGridView1.Columns.Add("serial_number", "Serial Number");
            //dataGridView1.Columns.Add("vendor_code", "Vendor Code");

            dataGridView1.GridColor = Color.Orange;
            dataGridView1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            // this.dataGridView1.Font = new Font("Goudy Stout", 8.25!,FontStyle.Regular, GraphicsUnit.Point, CType(0, Byte));
            dataGridView1.ColumnHeadersDefaultCellStyle.Font = new Font("Arial", 14F, FontStyle.Bold, GraphicsUnit.Pixel);
            dataGridView1.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.EnableResizing;
            this.dataGridView1.EnableHeadersVisualStyles = false;
            this.dataGridView1.ColumnHeadersHeight = 40;
            //foreach (DataColumn dc in dt.Columns)
            //{
            //    dataGridView1.Columns.Add(new DataGridViewTextBoxColumn());
            //}
            //foreach (DataRow dr in dt.Rows)
            //{
            //    dataGridView1.Rows.Add(dr.ItemArray);

            //}
            for (int j = 0; j < dataGridView1.Columns.Count; j++)
            {
                dataGridView1.Columns[j].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                dataGridView1.Columns[j].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;

            }

            //dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.Pink;


        }
        void bt_Click(object sender, EventArgs e)
        {
            saveData();
            setBarcodeSetting();
            PrintDocument pd = new PrintDocument();
            //Add PrintPage event handler
            //comboBox1.Text.ToString().Equals("")
            if (!prod_text.Text.ToString().Equals("") && !textBox2.Text.ToString().Equals("0"))
            {
                pd.PrintPage += new PrintPageEventHandler(CreateBarcodeSticker);
                for (int i = 0; i < Int32.Parse(textBox2.Text.ToString()); i++)
                {


                    preview_bool = true;
                    //adapter = new OleDbDataAdapter("select * from PageSetup", connection);
                    //ds = new DataSet();//student-> table name in stud.accdb file
                    //adapter.Fill(ds, "student");
                    //ds.Tables[0].Constraints.Add("pk_sno", ds.Tables[0].Columns[0], true);
                    //string str_size = ds.Tables[0].Rows[rno][1].ToString();
                    //MessageBox.Show(str_size);
                    float size = float.Parse("1.18");
                    size = (size + 0.005f) * 100;
                    //MessageBox.Show(size+"");
                    if (pd.PrinterSettings.PrinterName.CompareTo("TSC TTP-244 Pro") == 0)
                    {
                        pd.DefaultPageSettings.PaperSize = new System.Drawing.Printing.PaperSize("USER", 354, (int)size);
                        //Margins margins = new Margins(100,100,100,100);
                        // pd.DefaultPageSettings.Margins = margins;
                        //PrintPreviewDialog printPrvDlg = new PrintPreviewDialog();
                        //printPrvDlg.Document = pd;
                        //printPrvDlg.ShowDialog();

                        pd.Print();
                    }
                    else
                    {
                        MessageBox.Show("Default printer not set to TSC TTP-244 Pro. Please check Devices and Printers.");
                        break;
                    }
                    serial_nos++;
                }
                updateSerialNumber();
            }
            else
            {
                MessageBox.Show("Please select Product code or enter count of print");
            }
        }
        public void saveData()
        {
            try
            {
                //comboBox1.Text
                String query = "select * from mahindra_barcode WHERE `product_code`='" + prod_text.Text + "' and `vendor_code`='" + comboBox2.Text + "'";
                adapter = new OleDbDataAdapter(query, connection);
                ds = new DataSet();//student-> table name in stud.accdb file
                adapter.Fill(ds, "mahindra_barcode");
                ds.Tables[0].Constraints.Add("pk_sno", ds.Tables[0].Columns[0], true);
                int op_count = ds.Tables[0].Rows.Count;
                if (op_count < 1)
                {

                    MessageBox.Show("Please Enter valid product code. Product code does not exist in system please contact Administration.");
                    // connection.Open();
                    //OleDbCommand command = new OleDbCommand();
                    //command.Connection = connection;
                    //DateTime theDate;
                    //theDate = DateTime.Now;
                    //string d = theDate.ToString();
                    //string today_date = theDate.ToString("yyyy-MM-dd");
                    //command.CommandText = "insert into mahindra_barcode(`product_code`,`vendor_code`) values('" + comboBox1.Text + "','" + comboBox2.Text + "')";
                    //command.ExecuteNonQuery();
                    //MessageBox.Show("Data Saved" + comboBox1.Text);

                    //connection.Close();
                    //product_code_list.Add(comboBox1.Text.ToString());
                    //vendor_code.Add(comboBox2.Text.ToString());
                    //comboBox1.Items.Add(comboBox1.Text.ToString());
                    //comboBox2.Items.Add(comboBox2.Text.ToString());
                }
            }
            catch (Exception Ex)
            {
                MessageBox.Show("Error" + Ex);

            }
        }
        public void updateSerialNumber()
        {
            connection.Open();
            OleDbCommand command = new OleDbCommand();
            command.Connection = connection;

            command.CommandText = "Update print_setting set `serial_number`='" + serial_nos + "'";
            command.ExecuteNonQuery();
            connection.Close();
        }
        public String getBarcodeString()
        {
            String barcode_string = "";

            DateTime theDate = DateTime.Now;

            if (bar_product)
            {
                barcode_string = barcode_string + "" + prod_text.Text.ToString();
            }
            if (bar_rev_level)
            {
                barcode_string = barcode_string + "" + textBox3.Text.ToString();
            }
            if (bar_vendor)
            {
                barcode_string = barcode_string + "" + comboBox2.Text.ToString();
            }

           

            if (bar_date)
            {
                barcode_string = barcode_string + "" + theDate.ToString(datefmt);

            }
            if (bar_serial)
            {
                barcode_string = barcode_string + "" + serial_nos.ToString().PadLeft(6, '0');
            }
         
            return barcode_string;
        }
        public void CreateBarcodeSticker(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            //this prints the reciept
            Graphics graphic = e.Graphics;

            Font font = new Font("Arial", 7.7f, FontStyle.Bold); //must use a mono spaced font as the spaces need to line up
            float fontHeight = font.GetHeight();
            int startX = 5;
            int startY = 15;
            int offset = 0;
            DateTime theDate;
            theDate = DateTime.Now;

            //comboBox1.Text;
            string part_num = prod_text.Text;
            //string leakage = comboBox2.Text;
            string date_st = theDate.ToString(datefmt);
            String displaydate = theDate.ToString(display_date_fmt);
           
            MessagingToolkit.QRCode.Codec.QRCodeEncoder encoder = new MessagingToolkit.QRCode.Codec.QRCodeEncoder();
            encoder.QRCodeScale = 8;
            Bitmap btmp = encoder.Encode(getBarcodeString());

            graphic.DrawImage(btmp, startX +20, startY, 70, 70);
            graphic.DrawImage(btmp, startX + 255, startY, 70, 70);
            startX = startX + 95;
            if (display_product)
            {
                //comboBox1.Text
                graphic.DrawString(prod_text.Text.ToString(), font, new SolidBrush(Color.Black), startX, startY);
                // string top = "Item Name".PadRight(30) + "Price";
                offset = offset + (int)fontHeight + 1;
            }
            if (display_rev_level)
            {
                graphic.DrawString(textBox3.Text.ToString(), font, new SolidBrush(Color.Black), startX, startY + offset);
                offset = offset + (int)fontHeight + 1;
            }
            if (display_vendor)
            {
                graphic.DrawString(comboBox2.Text.ToString(), font, new SolidBrush(Color.Black), startX, startY + offset);
                offset = offset + (int)fontHeight + 1;
            }
            if (bar_date && display_serial)
            {
                graphic.DrawString(date_st + "" + serial_nos.ToString().PadLeft(6, '0'), font, new SolidBrush(Color.Black), startX, startY + offset);
                offset = offset + (int)fontHeight + 1;
            }
            else if (display_serial)
            {
                graphic.DrawString(serial_nos.ToString().PadLeft(6, '0'), font, new SolidBrush(Color.Black), startX, startY + offset);
                offset = offset + (int)fontHeight + 1;
            }
            else if (bar_date)
            {
                graphic.DrawString(date_st, font, new SolidBrush(Color.Black), startX, startY + offset);
                offset = offset + (int)fontHeight + 1;
            }
            if (display_descripton)
            {
                graphic.DrawString(textBox1.Text.ToString(), font, new SolidBrush(Color.Black), startX, startY + offset);
                offset = offset + (int)fontHeight + 1;
            }
            if (display_vailp_rev)
            {
                graphic.DrawString(textBox4.Text.ToString().PadRight(10,' ') +" "+theDate.ToString("ddMMyy hh:mm tt"), font, new SolidBrush(Color.Black), startX, startY + offset);
                offset = offset + (int)fontHeight + 1;
            }

           
          //  String query = "Insert into report(`barcode`,`serial_number`,`product_code`,`date_entry`,`time_entry`,`vendor_code`,`dt`,`count`,`descrip`) values('" +
             //   getBarcodeString() + "','" + serial_nos + "','" + comboBox1.Text.ToString() + "','" + theDate.ToString("dd/MM/yyyy") + "','" +
               // theDate.ToString("hh:mm:ss tt") + "','" + comboBox2.Text.ToString() + "','" + theDate.ToString("yyyy-MM-dd") + "','" + textBox2.Text + "','" + textBox1.Text + "')";
            //if (preview_bool)
               // insertToDb(query);

        }
        public void refreshForm()
        {
            setAutocomplete();
        }
        private void insertToDb(string query,bool load_req=true)
        {
            try
            {
                connection.Open();
                OleDbCommand command = new OleDbCommand();
                command.Connection = connection;

                command.CommandText = query;
                command.ExecuteNonQuery();
                connection.Close();
                if (load_req)
                {
                    loadData();
                }
            }catch(Exception ex){
                LogError(ex.Message.ToString());
            }
        }

        public void setAutocomplete()
        {
            String query = "select * from mahindra_barcode";
            if (product_code_list != null)
            {
                comboBox1.Items.Clear();
            }
            product_code_list = new List<string>();
            vendor_code = new List<string>();
            adapter = new OleDbDataAdapter(query, connection);
            dt = new DataTable();//student-> table name in stud.accdb file
            adapter.Fill(dt);
            foreach (DataRow dataRow in dt.Rows)
            {
                product_code_list.Add(dataRow.Field<String>("product_code"));
                vendor_code.Add(dataRow.Field<String>("vendor_code"));
            }
            comboBox1.Items.AddRange(product_code_list.ToArray<String>());
            comboBox1.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            comboBox1.AutoCompleteSource = AutoCompleteSource.ListItems;
            comboBox2.Items.AddRange(vendor_code.ToArray<String>());
            comboBox2.AutoCompleteMode = AutoCompleteMode.Append;
            comboBox2.AutoCompleteSource = AutoCompleteSource.ListItems;

        }
        public void setBarcodeSetting()
        {
            String query = "select * from print_setting";
            adapter = new OleDbDataAdapter(query, connection);
            dt = new DataTable();//student-> table name in stud.accdb file
            adapter.Fill(dt);
            foreach (DataRow dataRow in dt.Rows)
            {

                display_product = getValueOf(dataRow.Field<Int32>("display_product"));
                display_date = getValueOf(dataRow.Field<Int32>("display_date"));
                display_vendor = getValueOf(dataRow.Field<Int32>("display_vendor"));
                display_descripton = getValueOf(dataRow.Field<Int32>("display_desc"));
                display_serial = getValueOf(dataRow.Field<Int32>("display_serial"));

                bar_product = getValueOf(dataRow.Field<Int32>("bar_product"));

                bar_date = getValueOf(dataRow.Field<Int32>("bar_date"));

                bar_vendor = getValueOf(dataRow.Field<Int32>("bar_vendor"));

                bar_serial = getValueOf(dataRow.Field<Int32>("bar_serial"));

                bar_rev_level = getValueOf(dataRow.Field<Int32>("bar_rev_level"));
                display_rev_level = getValueOf(dataRow.Field<Int32>("display_rev_level"));
                display_vailp_rev = getValueOf(dataRow.Field<Int32>("display_vaipl"));

                switch (dataRow.Field<Int32>("date_format"))
                {
                    case 1:
                        datefmt = "MMyy";
                        break;
                    case 2:
                        datefmt = "ddMMyy";
                        break;
                    case 3:
                        datefmt = "ddMMyyHHmm";
                        break;

                }
                switch (dataRow.Field<Int32>("dis_date_frmt"))
                {
                    case 1:
                        display_date_fmt = "MMyy";
                        break;
                    case 2:
                        display_date_fmt = "MM/yy";
                        break;
                    case 3:
                        display_date_fmt = "ddMMyy HH:mm";
                        break;

                }
                DateTime dat = DateTime.Now;
                int mon = Int32.Parse(dat.ToString("MM"));
                if (mon.Equals(dataRow.Field<Int32>("cur_month")))
                    serial_nos = dataRow.Field<Int32>("serial_number");
                else
                {
                    serial_nos = 1;
                    resetSerialNumber(mon);
                }
            }
        }

        private void resetSerialNumber(int mon)
        {
            connection.Open();
            OleDbCommand command = new OleDbCommand();
            command.Connection = connection;

            command.CommandText = "Update print_setting set `serial_number`='" + 1 + "',`cur_month`='" + mon + "'";
            command.ExecuteNonQuery();
            connection.Close();
        }
        public bool getValueOf(int val)
        {
            if (val == 1)
                return true;
            else
                return false;
        }
        public void ApplyoperatorData()
        {
            error_label.Text = "";
            adapter = new OleDbDataAdapter("Select * from mahindra_barcode WHERE `product_code`='" + prod_text.Text + "'", connection);

            dt = new DataTable();//student-> table name in stud.accdb file
            adapter.Fill(dt);
            int count = dt.Rows.Count;
            if (count == 1)
            {
                foreach (DataRow dataRow in dt.Rows)
                {
                    comboBox2.Text = dataRow.Field<String>("mm_code");
                    textBox1.Text = dataRow.Field<String>("vendor_code");
                   // textBox3.Text = dataRow.Field<String>("cust_rev");
                  //  textBox4.Text = dataRow.Field<String>("vaipl_part");
                    prod_text.Text = dataRow.Field<String>("product_code");
                }
                textBox2.Focus();
            }
            else
            {
                error_label.Text ="Please Enter valid product code. Product code does not match exist with records in system please contact Administration.";
                prod_text.Focus();
                //MessageBox.Show("Please Enter valid product code. Product code does not match exist with records in system please contact Administration.");
                 
            }
        }
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

            adapter = new OleDbDataAdapter("Select * from mahindra_barcode WHERE `product_code`='" + comboBox1.Text + "'", connection);

            dt = new DataTable();//student-> table name in stud.accdb file
            adapter.Fill(dt);
            foreach (DataRow dataRow in dt.Rows)
            {
                comboBox2.Text = dataRow.Field<String>("mm_code");
               // textBox1.Text = dataRow.Field<String>("descrip");
                textBox3.Text = dataRow.Field<String>("vendor_code");
               // textBox4.Text = dataRow.Field<String>("vaipl_part");
                prod_text.Text = dataRow.Field<String>("product_code");
            }
            textBox2.Focus();
        }
        private void LogError(String ex)
        {
            string message = string.Format("Time: {0}", DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt"));
            message += Environment.NewLine;
            message += "-----------------------------------------------------------";
            message += Environment.NewLine;
            message += string.Format("Message: {0}", ex);
            message += Environment.NewLine;
            //message += string.Format("StackTrace: {0}", ex.StackTrace);
            //message += Environment.NewLine;
            //message += string.Format("Source: {0}", ex.Source);
            //message += Environment.NewLine;
            //message += string.Format("TargetSite: {0}", ex.TargetSite.ToString());
            //message += Environment.NewLine;
            message += "-----------------------------------------------------------";
            message += Environment.NewLine;
            string path = Application.StartupPath + "/OpenErrorLog.txt";
            using (StreamWriter writer = new StreamWriter(path, true))
            {
                writer.WriteLine(message);
                writer.Close();
            }
        }
        private void Form2_Load(object sender, EventArgs e)
        {
            //comboBox1.Focus();
            prod_text.Focus();
        }

        private void tblPnlDataEntry_Paint(object sender, PaintEventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }
        void changeSerialnumber()
        {

        }
        public void deleteOldRecords(String date="")
        {
            if (date == "")
            {
                date = DateTime.Now.AddDays(-5).ToString("dd/MM/yyyy");
            }
          //  String deleteString = "DELETE FROM report WHERE date_entry < #" + date + "";

          //  String deleteString = "DELETE * FROM report WHERE (((report.date_entry)<#" + date + "#))";

            String deleteString = "DELETE * FROM report WHERE (report.date_entry)<Date()-3";

            //String deleteString = "DELETE * FROM report WHERE (report.ID) < 552531";
           
            insertToDb(deleteString,false);

        }
        private void settingToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void xButton2_Click(object sender, EventArgs e)
        {
           //saveData();
            setBarcodeSetting();
            PrintDocument pd = new PrintDocument();
            //Add PrintPage event handler
            //comboBox1.Text.ToString().Equals("")
            if (!prod_text.Text.ToString().Equals("") && !textBox2.Text.ToString().Equals("0"))
            {
                pd.PrintPage += new PrintPageEventHandler(CreateBarcodeSticker);
                for (int i = 0; i <1; i++)
                {
                    //TODO change
                    preview_bool = true;
                    //adapter = new OleDbDataAdapter("select * from PageSetup", connection);
                    //ds = new DataSet();//student-> table name in stud.accdb file
                    //adapter.Fill(ds, "student");
                    //ds.Tables[0].Constraints.Add("pk_sno", ds.Tables[0].Columns[0], true);
                    //string str_size = ds.Tables[0].Rows[rno][1].ToString();
                    //MessageBox.Show(str_size);
                    float size = float.Parse("1.18");
                    size = (size + 0.005f) * 100;
                    //MessageBox.Show(size+"");
                    //if (pd.PrinterSettings.PrinterName.CompareTo("TSC TTP-244 Pro") == 0)
                    //{
                    pd.DefaultPageSettings.PaperSize = new System.Drawing.Printing.PaperSize("User", 354, (int)size);
                    //Margins margins = new Margins(100,100,100,100);
                    // pd.DefaultPageSettings.Margins = margins;
                    PrintPreviewDialog printPrvDlg = new PrintPreviewDialog();
                    printPrvDlg.Document = pd;
                    ((ToolStrip)printPrvDlg.Controls[1]).Items.Remove(((ToolStripButton)((ToolStrip)printPrvDlg.Controls[1]).Items[0]));
                    printPrvDlg.ShowDialog();

                   //  pd.Print();
                    //}
                    //else
                    //{
                    //    MessageBox.Show("Default printer not set to TSC TTP-244 Pro. Please check Devices and Printers.");
                    //}
                 //   serial_nos++;
                }
                // updateSerialNumber();
            }
            else
            {
                MessageBox.Show("Please select Product code or enter count of print");
            }

        }

        private void manageProductCodeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            CreateProduct cp = new CreateProduct(this);
            PasswordForm passwordfrm = new PasswordForm(cp);

            passwordfrm.Show();
        }

        private void configurationToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SettingScreen settingForm = new SettingScreen();
            PasswordForm passwd = new PasswordForm(settingForm);
            passwd.Show();
        }

        private void historyToolStripMenuItem_Click(object sender, EventArgs e)
        {
            History hist = new History();
            hist.Show();
        }

        private void refreshToolStripMenuItem_Click(object sender, EventArgs e)
        {
            refreshForm();

        }

        private void label13_Click(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void textBox2_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                xButton1.PerformClick();
            }
        }

        private void Form2_KeyDown(object sender, KeyEventArgs e)
        {

        }
        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (keyData == (Keys.Control | Keys.P))
            {
                //  MessageBox.Show("What the Ctrl+F?");
                xButton2.PerformClick();
                return true;
            }
            if (keyData == (Keys.Control | Keys.Enter))
            {
                //  MessageBox.Show("What the Ctrl+F?");
                xButton1.PerformClick();
                return true;
            }
            if (keyData == (Keys.Control | Keys.A))
            {
                //  MessageBox.Show("What the Ctrl+F?");
                xButton3.PerformClick();
                return true;
            }
            return base.ProcessCmdKey(ref msg, keyData);
        }

        private void xButton3_Click(object sender, EventArgs e)
        {
            comboBox1.Text = "";
            prod_text.Text = "";
            comboBox2.Text = "";
            textBox1.Text = "";
            textBox4.Text = "";
            textBox3.Text = "";
            prod_text.Focus();
        }

        private void prod_text_Enter(object sender, EventArgs e)
        {

        }

        private void comboBox1_KeyUp(object sender, KeyEventArgs e)
        {

        }

        private void prod_text_KeyUp(object sender, KeyEventArgs e)
        {
           /* if (e.KeyCode == Keys.Enter || e.KeyCode == Keys.Tab)
            {
                ApplyoperatorData();
            }
            else
            {
                comboBox2.Text = "";
                textBox1.Text = "";
                textBox4.Text = "";
                textBox3.Text = "";
                error_label.Text = "";
            }*/
        }

        private void prod_text_TextChanged(object sender, EventArgs e)
        {
            
        }

        private void xButton4_Click(object sender, EventArgs e)
        {
            ApplyoperatorData();
        }

        private void prod_text_Validated(object sender, EventArgs e)
        {

        }

        private void prod_text_KeyDown(object sender, KeyEventArgs e)
        {
            //if (e.KeyCode == Keys.Tab)
            //{
               // ApplyoperatorData();
            //}
        }

        private void prod_text_KeyPress(object sender, KeyPressEventArgs e)
        {
          // if (e.KeyChar == (char)Keys.Tab)
            //{
             //   ApplyoperatorData();
           // }
        }

        private void prod_text_PreviewKeyDown(object sender, PreviewKeyDownEventArgs e)
        {
            if (e.KeyData == Keys.Tab)
            {
                ApplyoperatorData();
            }
        }
    }

}