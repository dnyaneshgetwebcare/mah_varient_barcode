using System;
using System.Data;
using System.Data.OleDb;
using System.Drawing;

using System.Windows.Forms;

namespace TableLayoutPanelSample
{
    public partial class CreateProduct : Form
    {
        OleDbDataAdapter adapter;
        DataTable dt;
        DataSet ds;
        String idval = "";
        private OleDbConnection connection = new OleDbConnection();
        private Form2 form2;
        public CreateProduct()
        {
            MessageBox.Show("Hello");
            InitializeComponent();
            connection.ConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Application.StartupPath + "/voss_3_dataReader.accdb;Persist Security Info=True";

            loadData();
            DataGridViewLinkColumn dataButton = new DataGridViewLinkColumn();
            dataButton.HeaderText = "Action";
            dataButton.Text = "Edit";
            dataButton.UseColumnTextForLinkValue = true;
            dataButton.DataPropertyName = "lnkColumn";
            dataButton.LinkBehavior = LinkBehavior.SystemDefault;
            dataGridView1.Columns.Add(dataButton);
        }

        public CreateProduct(Form2 form2)
        {
            // TODO: Complete member initialization
            this.form2 = form2;
            InitializeComponent();
            connection.ConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Application.StartupPath + "/voss_3_dataReader.accdb;Persist Security Info=True";

            loadData();
            DataGridViewLinkColumn dataButton = new DataGridViewLinkColumn();
            dataButton.HeaderText = "Action";
            dataButton.Text = "Edit";
            dataButton.UseColumnTextForLinkValue = true;
            dataButton.DataPropertyName = "lnkColumn";
            dataButton.LinkBehavior = LinkBehavior.SystemDefault;
            dataGridView1.Columns.Add(dataButton);
        }

        private void xButton2_Click(object sender, EventArgs e)
        {
            if (!part_id.Text.Equals("") && !part_nos.Text.Equals(""))
                saveData();
            else
                MessageBox.Show("Product code or Vendor code cannot be blank. ");
        }
        public void saveData()
        {
            try
            {
                String query = "Select * from mahindra_barcode WHERE `product_id`='" + part_id.Text + "'";
                adapter = new OleDbDataAdapter(query, connection);
                ds = new DataSet();//student-> table name in stud.accdb file
                adapter.Fill(ds, "mahindra_barcode");
                ds.Tables[0].Constraints.Add("pk_sno", ds.Tables[0].Columns[0], true);
                int op_count = ds.Tables[0].Rows.Count;
                String query_insert = "";
                if (op_count < 1)
                {

                    //MessageBox.Show("Please Enter valid product code. Product code does not exist in system please contact Administration.");
                    query_insert = "insert into mahindra_barcode(`product_id`,`vendor_code`,`mm_part_code`,`part_desc` ,`oth` ,`cust_rev`) values('"
                        + part_id.Text + "','" + vendor_code.Text + "','" + part_nos.Text + "','" + part_desc.Text + "','"+ other_feild.Text+ "','" + cust_rev.Text + "')";
                    connection.Open();

                    OleDbCommand command = new OleDbCommand();
                    command.Connection = connection;

                    command.CommandText = query_insert;
                    command.ExecuteNonQuery();
                    MessageBox.Show("Data Saved");
                    connection.Close();
                    part_id.Text = "";
                    part_nos.Text = "";
                    vendor_code.Text = "";
                    part_desc.Text = "";
                   
                    other_feild.Text = "";
                    cust_rev.Text = "";
                    loadData();
                }
                else
                {
                    MessageBox.Show("Please Enter valid product code. Product code does already exist in system please contact Administration.");
                    // query_insert = "Update mahindra_barcode set `product_id`='" + textBox1.Text + "',`vendor_code`='" + textBox2.Text + "'";
                }




            }
            catch (Exception Ex)
            {
                MessageBox.Show("Error" + Ex);

            }
        }
        public void loadData()
        {
            String query = "select id as `Sr no`,product_id as `Product ID`,vendor_code as `Vendor Code`,mm_part_code as `M&M Part No`," +
                "part_desc as `Part Desc`,cust_rev as `Part Revsion`, oth as `Other` from mahindra_barcode order by id desc";
            adapter = new OleDbDataAdapter(query, connection);
            dt = new DataTable();//student-> table name in stud.accdb file
            adapter.Fill(dt);

            dataGridView1.DataSource = dt;
            //dataGridView1.Columns.Add("srn_no", "Sr No");
            //dataGridView1.Columns.Add("date_entry", "Date");
            //dataGridView1.Columns.Add("barcode", "Barcode");
            //dataGridView1.Columns.Add("product_id", "Product Code");
            //dataGridView1.Columns.Add("serial_number", "Serial Number");
            //dataGridView1.Columns.Add("vendor_code", "Vendor Code");

            dataGridView1.GridColor = Color.Orange;
            dataGridView1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            // this.dataGridView1.Font = new Font("Goudy Stout", 8.25!,FontStyle.Regular, GraphicsUnit.Point, CType(0, Byte));
            dataGridView1.ColumnHeadersDefaultCellStyle.Font = new Font("Arial", 14F, FontStyle.Bold, GraphicsUnit.Pixel);
            dataGridView1.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.EnableResizing;
            this.dataGridView1.EnableHeadersVisualStyles = false;
            // this.dataGridView1.ColumnHeadersHeight = 40;
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
        private void xButton1_Click(object sender, EventArgs e)
        {
            this.Hide();
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            // DataGridViewCell dataGridCell = (DataGridViewCell)sender;

            foreach (DataGridViewCell dataGridCell in dataGridView1.SelectedCells)
            {
                DataGridViewRow dr = dataGridView1.Rows[dataGridCell.RowIndex];
                idval = dr.Cells[1].Value.ToString();
                part_id.Text = dr.Cells[2].Value.ToString();
                vendor_code.Text = dr.Cells[3].Value.ToString();
                part_nos.Text = dr.Cells[4].Value.ToString();
                part_desc.Text = dr.Cells[5].Value.ToString();
                cust_rev.Text = dr.Cells[6].Value.ToString();
                other_feild.Text = dr.Cells[7].Value.ToString();
                //  Mf_date_code.Text = dr.Cells[7].Value.ToString();
                //  sl_no.Text = dr.Cells[8].Value.ToString();
                xButton1.Visible = true;
            }
        }



        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            foreach (DataGridViewRow row in dataGridView1.SelectedRows)
            {
                idval = row.Cells[1].Value.ToString();
                part_id.Text = row.Cells[2].Value.ToString();
                part_nos.Text = row.Cells[3].Value.ToString();
                vendor_code.Text = row.Cells[4].Value.ToString();
               /* vaipl_part.Text = row.Cells[5].Value.ToString();
                cust_rev.Text = row.Cells[6].Value.ToString();*/
                // Mf_date_code.Text = row.Cells[7].Value.ToString();
                // sl_no.Text = row.Cells[8].Value.ToString();
                xButton1.Visible = true;
                xButton2.Visible = false;

            }
        }

        private void xButton1_Click_1(object sender, EventArgs e)
        {
            if (!part_id.Text.Equals("") && !part_nos.Text.Equals(""))
            {


                try
                {
                    //String query = "Select * from mahindra_barcode WHERE `product_id`='" + textBox1.Text + "'";
                    //adapter = new OleDbDataAdapter(query, connection);
                    //ds = new DataSet();//student-> table name in stud.accdb file
                    //adapter.Fill(ds, "mahindra_barcode");
                    //ds.Tables[0].Constraints.Add("pk_sno", ds.Tables[0].Columns[0], true);
                    //int op_count = ds.Tables[0].Rows.Count;
                    String query_insert = "";
                    query_insert = "Update mahindra_barcode set `product_id`='" + part_id.Text + "',`vendor_code`='" + vendor_code.Text +
                        "',`mm_part_code`='" + part_nos.Text + "',`part_desc`='" + part_desc.Text + "',`cust_rev`='" + cust_rev.Text + "' ,`oth`='" + other_feild.Text + "' where id=" + idval;
                    connection.Open();
                    OleDbCommand command = new OleDbCommand();
                    command.Connection = connection;

                    command.CommandText = query_insert;
                    command.ExecuteNonQuery();
                    MessageBox.Show("Data Saved");
                }
                catch (Exception ex)
                {
                    MessageBox.Show("" + ex);

                }
                finally
                {
                    connection.Close();
                }
                part_id.Text = "";
                part_nos.Text = "";
                vendor_code.Text = "";
                part_desc.Text = "";
                other_feild.Text = "";
                //Mf_date_code.Text = "";
                // sl_no.Text = "";
                cust_rev.Text = "";
                idval = "";
                xButton1.Visible = false;
                loadData();

            }
            else
            {
                MessageBox.Show("Product code or Vendor code cannot be blank. ");
            }
        }

        private void CreateProduct_FormClosed(object sender, FormClosedEventArgs e)
        {
            form2.refreshForm();
        }

        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                part_nos.Focus();
            }
        }

        private void textBox2_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                vendor_code.Focus();
            }
        }

        private void xButton3_Click(object sender, EventArgs e)
        {
            part_id.Text = "";
            part_nos.Text = "";
            vendor_code.Text = "";
            part_desc.Text = "";
            // Mf_date_code.Text = "";
            //sl_no.Text = "";
            cust_rev.Text = "";
            other_feild.Text = "";
            xButton1.Visible = false;
            xButton2.Visible = true;
        }

        private void textBox3_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                if (xButton2.Visible)
                    xButton2.PerformClick();
                else if (xButton1.Visible)
                    xButton1.PerformClick();
            }
        }

        private void CreateProduct_FormClosing(object sender, FormClosingEventArgs e)
        {

        }

        private void label3_Click(object sender, EventArgs e)
        {

        }
    }
}