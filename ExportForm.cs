using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace TableLayoutPanelSample
{
    public partial class ExportForm : Form
    {

        public ExportForm()
        {
            InitializeComponent();

        }
        public string SqlString { get; set; }
        DataTable dt__export = new DataTable();
        SaveFileDialog saveFileDialog1;
        bool complt = true;
        private OleDbConnection connection = new OleDbConnection();
        private void ExportForm_Load(object sender, EventArgs e)
        {
           
            LogError("form On load ");
            try
            {
                connection.ConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Application.StartupPath + "/dataReader.accdb;Persist Security Info=True";
                //MessageBox.Show(this.SqlString);
                LogError("connection done ");
                dt__export.Columns.Add("Sr no.", typeof(int));
                dt__export.Columns.Add("Date", typeof(string));
                dt__export.Columns.Add("Barcode", typeof(string));
                dt__export.Columns.Add("Product Code", typeof(string));
                dt__export.Columns.Add("Vendor Code", typeof(string));
                dt__export.Columns.Add("Serial Number", typeof(string));
                dt__export.Columns.Add("Count", typeof(string));
                LogError("export column done ");
                // backgroundWorker1.WorkerSupportsCancellation = true;
                saveFileDialog1 = new SaveFileDialog();
                saveFileDialog1.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments); 
                saveFileDialog1.Title = "Save as Excel File";
                saveFileDialog1.FileName = "";
                saveFileDialog1.Filter = "Excel Files(2003)|*.xls|Excel Files(2007)|*.xlsx";
                LogError("dialog Open ");
            }catch(Exception ex){
                LogError(ex.ToString());
            }
            if (saveFileDialog1.ShowDialog() != DialogResult.Cancel)
            {
                LogError("Ok Button clicked ");
                backgroundWorker1.WorkerReportsProgress = true;
                backgroundWorker1.WorkerSupportsCancellation = true;

                backgroundWorker1.RunWorkerAsync();
                
                LogError("Back work start command ");

            }
            else
            {
                LogError("Cancel Button Clicked");
                backgroundWorker1.WorkerSupportsCancellation = true;

            }
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
            string path = Application.StartupPath + "/ErrorExportLog.txt";
            using (StreamWriter writer = new StreamWriter(path, true))
            {
                writer.WriteLine(message);
                writer.Close();
            }
        }

        private void MyFunction()
        {

        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            LogError("Background Work started");
            int srno = 0;
            try
            {
                DataGridView datagridViewexport = dataGridView2;
                if (connection.State.ToString() != "Open")
                    connection.Open();
                // Console.WriteLine(SqlString);
                OleDbCommand ocmd = new OleDbCommand(SqlString, connection);
                OleDbDataReader reader = ocmd.ExecuteReader();
                //MessageBox.Show(state+"");
                dt__export.Rows.Clear();
                LogError("Sr no "+ srno);
                while (reader.Read())
                {
                    srno++;
                    dt__export.Rows.Add(srno, reader["date_entry"].ToString().Substring(0, 10), reader["barcode"], reader["product_code"], reader["vendor_code"], reader["serial_number"], reader["count"]);
                }
                LogError("Sr no at end"+srno);

                // adapter.Fill(ds, "student");
                // ds.Tables[0].Constraints.Add("pk_sno", ds.Tables[0].Columns[0], true);
                //creating primary key for Tables[0] in dataset
                if (datagridViewexport.InvokeRequired)
                {
                    datagridViewexport.Invoke(new MethodInvoker(delegate { datagridViewexport.DataSource = dt__export; }));
                }
                else
                {
                    datagridViewexport.DataSource = dt__export;
                }
                LogError("Export to excel");
                // Console.WriteLine(datagridViewexport.Columns.Count + "");
                Microsoft.Office.Interop.Excel.Application ExcelApp = new Excel.Application();
                ExcelApp.Application.Workbooks.Add(Type.Missing);
                ExcelApp.Columns.ColumnWidth = 20;
                for (int i = 1; i < datagridViewexport.Columns.Count + 1; i++)
                {
                    if (backgroundWorker1.CancellationPending == true)
                    {
                        e.Cancel = true;
                        return;
                    }
                    ExcelApp.Cells[1, i] = datagridViewexport.Columns[i - 1].HeaderText;
                }

                for (int i = 0; i < datagridViewexport.Rows.Count; i++)
                {
                    if (backgroundWorker1.CancellationPending == true)
                    {
                        e.Cancel = true;
                        return;
                    }
                    datagridViewexport.Rows[i].DefaultCellStyle.BackColor = Color.Red;
                    for (int j = 0; j < datagridViewexport.Columns.Count; j++)
                    {
                        if (backgroundWorker1.CancellationPending == true)
                        {
                            e.Cancel = true;
                            return;
                        }
                        ExcelApp.Cells[i + 2, j + 1] = datagridViewexport.Rows[i].Cells[j].Value.ToString();
                        int percentage = (i + 1) * 100 / datagridViewexport.Rows.Count;
                        backgroundWorker1.ReportProgress(percentage);
                    }
                }
                LogError("Export to excel done");
               // ExcelApp.ActiveWorkbook.SaveCopyAs(saveFileDialog1.FileName.ToString());
                String temp = saveFileDialog1.FileName.ToString();
                var misValue = Type.Missing;
                ExcelApp.ActiveWorkbook.SaveAs(Environment.SpecialFolder.MyDocuments+"csharp-Excel.xls", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                ExcelApp.ActiveWorkbook.Saved = true;
                ExcelApp.ActiveWorkbook.Close(true);
                complt = false;
                LogError("Complete " + complt);
                MessageBox.Show("Data Exported Successfully");

            }
            catch (Exception ex)
            {
                LogError(ex.ToString());
                MessageBox.Show(ex.ToString());
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
             
            LogError("Work Complete " + complt+" "+e.Error);
            if (e.Error != null)
            {
              //  MessageBox.Show(e.Error+"");
                complt = false;
                this.Close();
            }
            else
            {
                this.Close();
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
    }
}