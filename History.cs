using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
namespace TableLayoutPanelSample
{
    public partial class History : Form
    {
        DataTable dt = new DataTable();
        int temp = 0;
        private OleDbConnection connection = new OleDbConnection();
        int i = 0;
        int rowCount;
        private int PageSize = 10;
        private int CurrentPageIndex = 1;
        private int TotalPage = 0;
        String frm_date, to_date;
        string today_date;
        string d;
        Random rnd = new Random();
        String sql = null;
        bool complt = true;
        DateTime theDate;
        OleDbDataAdapter adapter;
        DataSet ds;
        OleDbCommand ocmd;
        public History()
        {
            InitializeComponent();
            connection.ConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Application.StartupPath + "/dataReader.accdb;Persist Security Info=True";

            WindowState = FormWindowState.Maximized;
        }
        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);
            connection.Open();
            dt.Columns.Add("Sr no.", typeof(int));
            dt.Columns.Add("Date", typeof(string));
            dt.Columns.Add("Barcode", typeof(string));
            dt.Columns.Add("Product Code", typeof(string));
            dt.Columns.Add("Vendor Code", typeof(string));
            dt.Columns.Add("Serial Number", typeof(string));
            dt.Columns.Add("Description", typeof(string));
            dt.Columns.Add("Count", typeof(string));
            backgroundWorker1.WorkerReportsProgress = true;
            backgroundWorker1.WorkerSupportsCancellation = true;

            // dt.Columns.Add("",typeof(string));
            loaddata();
        }
        void loaddata()
        {
            try
            {

                theDate = DateTime.Now;
                d = theDate.ToString();
                today_date = theDate.ToString("dd/MM/yyyy");
                //today_date = "2016-03-13";
                sql = "select * from report WHERE date_entry=#" + today_date + "# order by ID ASC";
                adapter = new OleDbDataAdapter(sql, connection);
                ds = new DataSet();//student-> table name in stud.accdb file
                adapter.Fill(ds, "student");
                /*dataGridView1.DataSource = ds;
                dataGridView1.DataMember = "student";*/
                this.CalculateTotalPages();
                this.dataGridView1.ReadOnly = true;
                this.dataGridView1.DataSource = GetCurrentRecords(1, connection);

                this.dataGridView1.GridColor = Color.Orange;
                // this.dataGridView1.Font = new Font("Goudy Stout", 8.25!,FontStyle.Regular, GraphicsUnit.Point, CType(0, Byte));
                this.dataGridView1.ColumnHeadersDefaultCellStyle.Font = new Font("Arial", 14F, FontStyle.Bold, GraphicsUnit.Pixel);
                this.dataGridView1.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.EnableResizing;
                this.dataGridView1.EnableHeadersVisualStyles = false;
                this.dataGridView1.ColumnHeadersHeight = 40;
                for (int j = 0; j < dataGridView1.Columns.Count; j++)
                {
                    this.dataGridView1.Columns[j].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;

                }
                //dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.Pink;

                //this.dataGridView1.Columns[0].Width = 30;
                //this.dataGridView1.Columns[1].Width = 50;
                //this.dataGridView1.Columns[2].Width = 40;
                //this.dataGridView1.Columns[3].Width = 40;
                //this.dataGridView1.Columns[5].Width = 30;
                //this.dataGridView1.Columns[4].Width = 40;
                //this.dataGridView1.Columns[6].Width = 70;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }


        }
        void loaddataRangeWise(string sdate, string edate)
        {
            temp = 1;
            frm_date = sdate;
            to_date = edate;
            sql = "select * from report WHERE date_entry BETWEEN #" + sdate + "# AND #" + edate + "# order by ID ASC";
         //   Console.WriteLine(sql);
            adapter = new OleDbDataAdapter(sql, connection);
            ds = new DataSet();//student-> table name in stud.accdb file
            adapter.Fill(ds, "student");
            /*dataGridView1.DataSource = ds;
            dataGridView1.DataMember = "student";*/
            this.CalculateTotalPages();
            this.dataGridView1.ReadOnly = true;
            this.dataGridView1.DataSource = GetCurrentRecords_range(1, connection, sdate, edate);
            //  adapter = new OleDbDataAdapter("select * from report WHERE date_entry BETWEEN '" + sdate + "' AND '" + edate + "' order by ID ASC", connection);
            //OleDbCommand ocmd = new OleDbCommand("select * from report WHERE date_entry BETWEEN '" + sdate + "' AND '" + edate + "' order by ID ASC", connection);

            // OleDbDataReader reader = ocmd.ExecuteReader();
            //MessageBox.Show(state+"");
            /*
                dt.Rows.Clear();
                while (reader.Read())
                {
                    i++;
                    //MessageBox.Show(reader["Operator Name"] + "");
                    /* firstName.Add(reader["FirstName"].ToString());
                     lastName.Add(reader["LastName"].ToString());*/
            /*   dt.Rows.Add(i, reader["Operator Name"], reader["Part Number"], reader["date_entry"], reader["Time"], reader["Leak"], reader["Result"]);
           }
            

           // adapter.Fill(ds, "student");
           // ds.Tables[0].Constraints.Add("pk_sno", ds.Tables[0].Columns[0], true);
           //creating primary key for Tables[0] in dataset
           this.dataGridView1.DataSource = dt;
           */
            // this.dataGridView1.Font = new Font("Goudy Stout", 8.25!,FontStyle.Regular, GraphicsUnit.Point, CType(0, Byte));
            this.dataGridView1.ColumnHeadersDefaultCellStyle.Font = new Font("Arial", 14F, FontStyle.Bold, GraphicsUnit.Pixel);
            this.dataGridView1.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.EnableResizing;
            this.dataGridView1.EnableHeadersVisualStyles = false;
            this.dataGridView1.ColumnHeadersHeight = 40;
            for (int j = 0; j < dataGridView1.Columns.Count; j++)
            {
                this.dataGridView1.Columns[j].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;

            }
            //dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.Pink;

        }
        private DataTable GetCurrentRecords_range(int page, OleDbConnection connection, string sdate, string edate)
        {
            OleDbDataReader reader;
            // int PreviouspageLimit;
            // ocmd = new OleDbCommand("select * from report WHERE date_entry='" + today_date + "' order by ID ASC", connection);

            // adapter.Fill(ds, "student");
            // ds.Tables[0].Constraints.Add("pk_sno", ds.Tables[0].Columns[0], true);
            //creating primary key for Tables[0] in dataset
            // today_date = "2016-03-13";
            if (page == 1)
            {
                i = 0;
                ocmd = new OleDbCommand("SELECT report.* FROM report WHERE report.ID In(SELECT TOP " + PageSize + " A.ID FROM [SELECT TOP " + PageSize + " report.ID FROM report where date_entry BETWEEN #" + sdate + "# AND #" + edate + "# ORDER BY report.ID]. AS A ORDER BY report.ID  DESC) ORDER BY report.ID", connection);
                //  ocmd = new OleDbCommand("select TOP " + PageSize + " * from report WHERE date_entry BETWEEN '" + sdate + "' AND '" + edate + "' order by ID ASC", connection);
                //Console.WriteLine("select TOP " + PageSize + " * from report WHERE date_entry BETWEEN '" + sdate + "' AND '" + edate + "' order by ID ASC");
                // cmd2 = new SqlCommand("Select TOP " + PageSize + " * from Customers ORDER BY CustomerID", con);
            }
            else
            {
                int PreviouspageLimit = (page - 1) * PageSize;
                i = PreviouspageLimit;
                int offset = PreviouspageLimit + PageSize;
                int temp_pagesize = PageSize;

                if (offset > rowCount)
                {
                    temp_pagesize = rowCount % PageSize;
                    if (temp_pagesize == 0)
                    {
                        temp_pagesize = PageSize;
                    }

                }

                ocmd = new OleDbCommand("SELECT report.* FROM report WHERE report.ID In(SELECT TOP " + temp_pagesize + " A.ID FROM [SELECT TOP " + offset + " report.ID FROM report where date_entry BETWEEN #" + sdate + "# AND #" + edate + "# ORDER BY report.ID]. AS A ORDER BY report.ID  DESC) ORDER BY report.ID", connection);
                // ocmd = new OleDbCommand("select TOP " + PageSize + " * from report WHERE date_entry BETWEEN '" + sdate + "' AND '" + edate + "' AND ID NOT IN (Select TOP " + PreviouspageLimit + " ID from report WHERE date_entry BETWEEN '" + sdate + "' AND '" + edate + "')", connection);
                Console.WriteLine("SELECT report.* FROM report WHERE report.ID In(SELECT TOP " + temp_pagesize + " A.ID FROM [SELECT TOP " + offset + " report.ID FROM report where date_entry BETWEEN #" + sdate + "# AND #" + edate + "# ORDER BY report.ID]. AS A ORDER BY report.ID  DESC) ORDER BY report.ID");

            }
            try
            {
                // con.Open();
                reader = ocmd.ExecuteReader();

                dt.Rows.Clear();
                while (reader.Read())
                {
                    i++;
                    String temp = reader["ID"].ToString();
                    Console.WriteLine(i + " " + temp);
                    //MessageBox.Show(reader["Operator Name"] + "");
                    /* firstName.Add(reader["FirstName"].ToString());
                     lastName.Add(reader["LastName"].ToString());*/
                    // Console.WriteLine(dt.Rows.Count+" col "+dt.Columns.Count);
                    dt.Rows.Add(i, reader["date_entry"].ToString().Substring(0, 10), reader["barcode"], reader["product_code"], reader["vendor_code"], reader["serial_number"], reader["descrip"], reader["count"]);
                }
                //MessageBox.Show(state+"");

                /* this.adp1.SelectCommand = cmd2;
                 this.adp1.Fill(dt);*/
            }
            catch (Exception e)
            {
                Debug.Print(e.ToString());
                Console.WriteLine(e);
            }



            return dt;
        }
        private void CalculateTotalPages()
        {
            rowCount = ds.Tables["student"].Rows.Count;
            if (label4.InvokeRequired)
            {
                label4.Invoke(new MethodInvoker(delegate { label4.Text = "No of records " + rowCount; }));
            }
            else
            {
                label4.Text = "No of records " + rowCount;
            }
            this.TotalPage = rowCount / PageSize;
            if (rowCount % PageSize > 0) // if remainder is more than  zero 
            {
                this.TotalPage += 1;
            }
        }
        private DataTable GetCurrentRecords(int page, OleDbConnection con)
        {

            OleDbDataReader reader;

            if (page == 1)
            {
                i = 0;
                ocmd = new OleDbCommand("SELECT report.* FROM report WHERE report.ID In(SELECT TOP " + PageSize + " A.ID FROM [SELECT TOP " + PageSize + " report.ID FROM report where date_entry =#" + today_date + "# ORDER BY report.ID]. AS A ORDER BY report.ID  DESC) ORDER BY report.ID", connection);

            }
            else
            {
                int PreviouspageLimit = (page - 1) * PageSize;
                i = PreviouspageLimit;
                int offset = PreviouspageLimit + PageSize;
                int temp_pagesize = PageSize;

                if (offset > rowCount)
                {
                    temp_pagesize = rowCount % PageSize;
                    if (temp_pagesize == 0)
                    {
                        temp_pagesize = PageSize;
                    }

                }

                ocmd = new OleDbCommand("SELECT report.* FROM report WHERE report.ID In(SELECT TOP " + temp_pagesize + " A.ID FROM [SELECT TOP " + offset + " report.ID FROM report where date_entry =#" + today_date + "# ORDER BY report.ID]. AS A ORDER BY report.ID  DESC) ORDER BY report.ID", connection);

            }
            try
            {
                reader = ocmd.ExecuteReader();

                dt.Rows.Clear();
                while (reader.Read())
                {
                    i++;

                    dt.Rows.Add(i, reader["date_entry"].ToString().Substring(0, 10), reader["barcode"], reader["product_code"], reader["vendor_code"], reader["serial_number"], reader["descrip"], reader["count"]);
                }

            }
            catch (Exception e)
            {
                Debug.Print(e.ToString());
                Console.WriteLine(e);
            }



            return dt;
        }

        private void search_btn_Click(object sender, EventArgs e)
        {
            string StartDate = dateTimePicker1.Value.ToString("dd/MM/yyyy");
            string EndDate = dateTimePicker2.Value.ToString("dd/MM/yyyy");
            loaddataRangeWise(StartDate, EndDate);
        }
        public DataTable getExportData()
        {
            DataTable resultTable = new DataTable();
            String query = sql;
            resultTable.Columns.Add("Sr no.", typeof(int));
            resultTable.Columns.Add("Date", typeof(string));
            resultTable.Columns.Add("Barcode", typeof(string));
            resultTable.Columns.Add("Product Code", typeof(string));
            resultTable.Columns.Add("Vendor Code", typeof(string));
            resultTable.Columns.Add("Serial Number", typeof(string));
            resultTable.Columns.Add("Description", typeof(string));
            resultTable.Columns.Add("Count", typeof(string));
            OleDbDataReader reader;

            try
            {

               // connection.Open();
                OleDbCommand cmd = new OleDbCommand();
                cmd.CommandText = query;
                cmd.Connection = connection;
                reader = cmd.ExecuteReader();
                resultTable.Rows.Clear();
                int srnos = 0;
                while (reader.Read())
                {
                    srnos++;
                    resultTable.Rows.Add(srnos, reader["date_entry"].ToString().Substring(0, 10), reader["barcode"], reader["product_code"], reader["vendor_code"], reader["serial_number"], reader["descrip"], reader["count"]);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {


          //      connection.Close();
            }
            return resultTable;
        }
        private void export_btn_Click(object sender, EventArgs e)
        {
            tableLayoutPanel5.Visible = true;
            backgroundWorker1.WorkerReportsProgress = true;
            backgroundWorker1.WorkerSupportsCancellation = true;

            backgroundWorker1.RunWorkerAsync();
            //LogError("Enter to export");
            //ExportForm from6 = new ExportForm();
            //from6.SqlString = sql;
            //LogError("sql string ::"+sql);
            //from6.Show();
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
            string path = Application.StartupPath +"/ErrorLog.txt";
            using (StreamWriter writer = new StreamWriter(path, true))
            {
                writer.WriteLine(message);
                writer.Close();
            }
        }
        private void first_page_btn_Click(object sender, EventArgs e)
        {
            this.CurrentPageIndex = 1;
            if (temp == 0)
                this.dataGridView1.DataSource = GetCurrentRecords(this.CurrentPageIndex, connection);
            else
                this.dataGridView1.DataSource = GetCurrentRecords_range(this.CurrentPageIndex, connection, frm_date, to_date);
        }

        private void previous_btn_Click(object sender, EventArgs e)
        {
            if (this.CurrentPageIndex > 1)
            {
                this.CurrentPageIndex--;
                if (temp == 0)
                    this.dataGridView1.DataSource = GetCurrentRecords(this.CurrentPageIndex, connection);
                else
                    this.dataGridView1.DataSource = GetCurrentRecords_range(this.CurrentPageIndex, connection, frm_date, to_date);
            }
        }

        private void next_btn_Click(object sender, EventArgs e)
        {
            if (this.CurrentPageIndex < this.TotalPage)
            {
                this.CurrentPageIndex++;
                if (temp == 0)
                    this.dataGridView1.DataSource = GetCurrentRecords(this.CurrentPageIndex, connection);
                else
                    this.dataGridView1.DataSource = GetCurrentRecords_range(this.CurrentPageIndex, connection, frm_date, to_date);
            }
        }
        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            LogError("Background Work started");
            int srno = 0;
            try
            {
                //DataGridView datagridViewexport = dataGridView1;
                DataTable exportDataTable = getExportData();
                Microsoft.Office.Interop.Excel.Application ExcelApp = new Excel.Application();
                ExcelApp.Application.Workbooks.Add(Type.Missing);
                ExcelApp.Columns.ColumnWidth = 20;
                for (int i = 1; i < exportDataTable.Columns.Count + 1; i++)
                {
                    if (backgroundWorker1.CancellationPending == true)
                    {
                        e.Cancel = true;
                        return;
                    }
                    ExcelApp.Cells[1, i] = exportDataTable.Columns[i - 1].ColumnName;
                }

                for (int i = 0; i < exportDataTable.Rows.Count; i++)
                {
                    if (backgroundWorker1.CancellationPending == true)
                    {
                        e.Cancel = true;
                        return;
                    }
                   // datagridViewexport.Rows[i].DefaultCellStyle.BackColor = Color.Red;
                    for (int j = 0; j < exportDataTable.Columns.Count; j++)
                    {
                        if (backgroundWorker1.CancellationPending == true)
                        {
                            e.Cancel = true;
                            return;
                        }
                        if (j == 2)
                        {
                            ExcelApp.Cells[i + 2, j + 1] = "'" + exportDataTable.Rows[i][j].ToString();
                        }
                        else
                        {
                            ExcelApp.Cells[i + 2, j + 1] = exportDataTable.Rows[i][j].ToString();

                        }
                        int percentage = (i + 1) * 100 / exportDataTable.Rows.Count;
                        backgroundWorker1.ReportProgress(percentage);
                    }
                }
                LogError("Export to excel done");
                // ExcelApp.ActiveWorkbook.SaveCopyAs(saveFileDialog1.FileName.ToString());
                String tem = DateTime.Now.ToString("dd_MM_yy");
                String file_name = "Print_record" + tem + "_" + rnd.Next(1000) + ".xls";
                var misValue = Type.Missing;
                ExcelApp.ActiveWorkbook.SaveAs( file_name, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                ExcelApp.ActiveWorkbook.Saved = true;
                ExcelApp.ActiveWorkbook.Close(true);
                complt = false;
                LogError("Complete " + complt);
                MessageBox.Show("Data Exported to "+file_name+" Successfully");
               
                
            }
            catch (Exception ex)
            {
                LogError(ex.ToString());
                MessageBox.Show("Fail to export Please try again.");
            }
        }

        private void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progressBar1.Value = e.ProgressPercentage;

            if (label1.InvokeRequired)
            {
                label1.Invoke(new MethodInvoker(delegate { label1.Text = e.ProgressPercentage.ToString() + "%"; }));
            }
            else
            {
                label1.Text = e.ProgressPercentage.ToString() + "%";
            }
        }

        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {

            LogError("Work Complete " + complt + " " + e.Error);
            if (e.Error != null)
            {
                //  MessageBox.Show(e.Error+"");
                complt = false;
               
               // this.Close();
            }
            progressBar1.Value = 100;
            if (label1.InvokeRequired)
            {
                label1.Invoke(new MethodInvoker(delegate { label1.Text = 100 + "%"; }));
            }
            else
            {
                label1.Text = 100 + "%";
            }

        }
        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            LogError("Form Closing " + complt);
            base.OnFormClosing(e);
            //this.Visible = false;
            if (complt)
            {
                if (e.CloseReason == CloseReason.WindowsShutDown) return;

                // Confirm user wants to close
                switch (MessageBox.Show(this, "Are you sure you want to close?", "Closing", MessageBoxButtons.YesNo))
                {
                    case DialogResult.No:
                        e.Cancel = true;
                        break;
                    default:
                        backgroundWorker1.CancelAsync();
                        break;
                }
            }
            else
            {
                try
                {
                    this.Hide();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex + "");
                }
            }
        }
        private void last_page_btn_Click(object sender, EventArgs e)
        {
            this.CurrentPageIndex = TotalPage;
            if (temp == 0)
                this.dataGridView1.DataSource = GetCurrentRecords(this.CurrentPageIndex, connection);
            else
                this.dataGridView1.DataSource = GetCurrentRecords_range(this.CurrentPageIndex, connection, frm_date, to_date);
        }

        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }

}