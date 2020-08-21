using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SQLite;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Net;
using System.Globalization;
using System.IO;

namespace WF_JGB_Invoices_001
{
    public partial class Form1 : Form
    {
        #region FIELDS

        frmProgress progressDlg;

        #endregion

        public string csvFilePath = "invoice.csv";

        public Form1()
        {
            InitializeComponent();
        }

        private SQLiteConnection CreateConnection()
        {

            SQLiteConnection sqlite_conn;
            // Create a new database connection:
            sqlite_conn = new SQLiteConnection("Data Source=111.db; Version = 3; New = True; Compress = True; ");
            // Open the connection:
            try
            {
                sqlite_conn.Open();
            }
            catch (Exception ex)
            {

            }
            return sqlite_conn;
        }

        private void CreateTable(SQLiteConnection conn)
        {

            SQLiteCommand sqlite_cmd;
            string Createsql = "CREATE TABLE tblStatus(Col1 VARCHAR(20), Col2 INT)";
            string Createsql1 = "CREATE TABLE SampleTable1(Col1 VARCHAR(20), Col2 INT)";
            sqlite_cmd = conn.CreateCommand();
            sqlite_cmd.CommandText = Createsql;
            sqlite_cmd.ExecuteNonQuery();
            sqlite_cmd.CommandText = Createsql1;
            sqlite_cmd.ExecuteNonQuery();

        }

        static void InsertSubmitStatusData(SQLiteConnection conn, string invName, string invNum, string invDate, string custPo, string responseTime, string responseStatus)
        {
            SQLiteCommand sqlite_cmd;
            sqlite_cmd = conn.CreateCommand();
            sqlite_cmd.CommandText = "INSERT INTO tblSubmitStatus(ShipToAdd1, InvNum, InvDate, CustPo, ResponseTime, ResponseStatus)" +
                " VALUES('" + invName + "', '" + invNum + "', '" + invDate + "', '" + custPo + "', '" + responseTime + "', '" + responseStatus + "'); ";
            sqlite_cmd.ExecuteNonQuery();
            conn.Close();

        }

        private void ShowInvoice(SQLiteConnection conn, string name)
        {
            SQLiteDataReader sqlite_datareader;
            SQLiteCommand sqlite_cmd;
            sqlite_cmd = conn.CreateCommand();

            if (name == "All")
            {
                sqlite_cmd.CommandText = "SELECT * FROM tblTest00";
            }
            else
            {
                sqlite_cmd.CommandText = "SELECT * FROM tblTest00 WHERE ShipToAdd1 LIKE '%" + name + "%'";
            }

            sqlite_datareader = sqlite_cmd.ExecuteReader();

            this.listView1.Items.Clear();
            while (sqlite_datareader.Read())
            {
                string invNum = sqlite_datareader.GetString(0);
                string invDate = sqlite_datareader.GetString(1);
                string custPo = sqlite_datareader.GetString(9);
                string saleOrder = Convert.ToString(sqlite_datareader.GetValue(sqlite_datareader.GetOrdinal("SupplierOrderNumber")));
                string shipNum = Convert.ToString(sqlite_datareader.GetValue(sqlite_datareader.GetOrdinal("ShipmentNum")));
                string shipDate = Convert.ToString(sqlite_datareader.GetValue(sqlite_datareader.GetOrdinal("ShipDate")));
                string invAmt = Convert.ToString(sqlite_datareader.GetValue(sqlite_datareader.GetOrdinal("ShipDate")));

                ListViewItem item;

                item = this.listView1.Items.Add(Convert.ToString(invNum));
                item.SubItems.Add(Convert.ToString(invDate));
                item.SubItems.Add(Convert.ToString(custPo));
                item.SubItems.Add(Convert.ToString(saleOrder));
                item.SubItems.Add(Convert.ToString(shipNum));
                item.SubItems.Add(Convert.ToString(shipDate));
                item.SubItems.Add("$" + Convert.ToString(invNum));
            }
            conn.Close();
        }

        private void ShowSubmitStatus(SQLiteConnection conn, string name)
        {
            SQLiteDataReader sqlite_datareader;
            SQLiteCommand sqlite_cmd;
            sqlite_cmd = conn.CreateCommand();

            if (name == "All")
            {
                sqlite_cmd.CommandText = "SELECT * FROM tblSubmitStatus";
            }
            else
            {
                sqlite_cmd.CommandText = "SELECT * FROM tblSubmitStatus WHERE ShipToAdd1 LIKE '%" + name + "%'";
            }

            sqlite_datareader = sqlite_cmd.ExecuteReader();

            this.listView2.Items.Clear();
            while (sqlite_datareader.Read())
            {
                string invNum = sqlite_datareader.GetString(1);
                string invDate = sqlite_datareader.GetString(2);
                string custPo = sqlite_datareader.GetString(3);
                string responseTime = sqlite_datareader.GetString(4);
                string responseStatus = sqlite_datareader.GetString(5);

                ListViewItem item;

                item = this.listView2.Items.Add(Convert.ToString(invNum));
                item.SubItems.Add(Convert.ToString(invDate));
                item.SubItems.Add(Convert.ToString(custPo));
                item.SubItems.Add(Convert.ToString(responseTime));
                item.SubItems.Add(Convert.ToString(responseStatus));
            }
            conn.Close();
        }

        private string FixDecimals(string strValue)
        {
            string retStr = "";

            if (!strValue.Contains("."))
            {
                return strValue;
            }

            double OutVal;
            double.TryParse(strValue, out OutVal);

            if (OutVal == 0 && strValue != "0")
            {
                return strValue;
            }

            if (double.IsNaN(OutVal) || double.IsInfinity(OutVal))
            {
                return strValue;
            }

            string[] parts = strValue.Split('.');
            int i1 = int.Parse(parts[0]);
            int i2 = int.Parse(parts[1]);

            if (i2 < 100)
            {
                retStr = string.Format("{0:N2}", strValue);
            }
            else
            {
                retStr = strValue;
            }
            return retStr;
        }

        private void DumpTableToCsv(SQLiteConnection conn, string tblName)
        {

            if (backgroundWorker2.IsBusy != true)
            {
                // create a new instance of the alert form
                progressDlg = new frmProgress();
                progressDlg.StartPosition = FormStartPosition.CenterParent;

                // event handler for the Cancel button in AlertForm
                progressDlg.Canceled += new EventHandler<EventArgs>(buttonCancel_Click);
                progressDlg.Show(this);
                // Start the asynchronous operation.
                backgroundWorker2.RunWorkerAsync();
            }

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            this.lblTitle.Text = "";
            this.lblTitle2.Text = "";
            this.tabControl1.Visible = false;

            listView1.Columns.Add("Inv Num", 100);
            listView1.Columns.Add("Inv Date", 100);
            listView1.Columns.Add("Cust PO", 100);
            listView1.Columns.Add("Sales Order", 100);
            listView1.Columns.Add("Ship Num", 100);
            listView1.Columns.Add("Ship Date", 100);
            listView1.Columns.Add("Inv Amt", 100);

            listView2.Columns.Add("Inv Num", 100);
            listView2.Columns.Add("Inv Date", 100);
            listView2.Columns.Add("Cust PO", 100);
            listView2.Columns.Add("Response Time", 120);
            listView2.Columns.Add("Response", 250);
            //CreateTable(sqlite_conn);
        }

        private void getListView1(string name)
        {
            SQLiteConnection sqlite_conn;
            sqlite_conn = CreateConnection();
            ShowInvoice(sqlite_conn, name);

            for (int i = 0; i <= listView1.Items.Count - 1; i = (i + 2))
            {
                listView1.Items[i].BackColor = Color.AliceBlue;
            }
            for (int i = 0; i <= listView1.Items.Count - 1; i++)
            {
                listView1.Items[i].Checked = true;
            }
        }

        private void getListView2(string name)
        {

            SQLiteConnection sqlite_conn;
            sqlite_conn = CreateConnection();
            ShowSubmitStatus(sqlite_conn, name);

            for (int i = 0; i <= listView1.Items.Count - 1; i = (i + 2))
            {
                listView1.Items[i].BackColor = Color.AliceBlue;
            }
        }

        private void ListView1_ItemChecked(object sender, ItemCheckedEventArgs e)
        {
            //ListViewItem l = listView1.Items[e.Index];
            if (e.Item.Checked)
                e.Item.Selected = true;

            if (!e.Item.Checked)
                e.Item.Selected = false;


        }

        private void ListView1_ItemSelectionChanged(object sender, ListViewItemSelectionChangedEventArgs e)
        {
            int i = 0;
            double amt = 0.0;
            foreach (ListViewItem item in listView1.Items)
            {
                if (item.Checked == true)
                {
                    amt += Convert.ToDouble((item.SubItems[6].Text).Replace("$", ""));
                    i++;
                }
            }
            this.label7.Text = i.ToString();
            this.label8.Text = string.Format("{0:C1}", amt);
        }

        private void WriteToCsvFile(DataTable dataTable, string filePath)
        {
            StringBuilder fileContent = new StringBuilder();

            foreach (var col in dataTable.Columns)
            {
                fileContent.Append(col.ToString() + "|");
            }

            fileContent.Replace("|", System.Environment.NewLine, fileContent.Length - 1, 1);

            foreach (DataRow dr in dataTable.Rows)
            {
                foreach (var column in dr.ItemArray)
                {
                    fileContent.Append(column.ToString() + "|");
                }

                fileContent.Replace("|", System.Environment.NewLine, fileContent.Length - 1, 1);
            }

            System.IO.File.WriteAllText(filePath, fileContent.ToString());
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            int i = 0;
            double amt = 0.0;

            // Convert the list view data to csv file
            SQLiteConnection sqlite_conn;
            sqlite_conn = CreateConnection();
            DumpTableToCsv(sqlite_conn, "tblTest00");

            int count = 0;
            while(true)
            {
                count++;
                csvFilePath = "invoice_" + Convert.ToString(count) + ".csv";
                if (File.Exists(csvFilePath))
                {
                    continue;
                }
                break;
            }

            //Upload the csv file to the ftp server
            string retStr = "";
            try
            {
                using (var client = new WebClient())
                {
                    client.Credentials = new NetworkCredential("test", "test$$$");
                    Byte[] ret = new Byte[1024];
                    ret = client.UploadFile("ftp://158.69.253.232:21/DX014/" + csvFilePath, WebRequestMethods.Ftp.UploadFile, csvFilePath);
                    retStr = Convert.ToString(ret);
                    retStr = "Success";
                }
            }
            catch (Exception ex)
            {
                retStr = ex.Message;
            }

            //Write the result to the SubmitStatus table in the database(SQLite)
            foreach (ListViewItem item in listView1.Items)
            {
                if (item.Checked == true)
                {
                    string invName = this.comboBox1.Text;
                    string invNum = Convert.ToString(item.SubItems[0].Text);
                    string invDate = Convert.ToString(item.SubItems[1].Text);
                    string custPo = Convert.ToString(item.SubItems[2].Text);
                    string responseTime = Convert.ToString(DateTimeOffset.UtcNow.ToUnixTimeSeconds());
                    string responseStatus = retStr;

                    SQLiteConnection sqlite_conn_01;
                    sqlite_conn_01 = CreateConnection();
                    InsertSubmitStatusData(sqlite_conn_01, invName, invNum, invDate, custPo, responseTime, responseStatus);
                    break;
                }
            }

            this.label7.Text = i.ToString();
            this.label8.Text = string.Format("{0:C1}", amt);

            this.tabControl1.SelectedTab = tabPage2;
            getListView2(this.comboBox1.Text);
            this.textBox5.Text = "HTTPS return string; Invoice 1036678 Data Received; confirmation Time 11/14/2017 12:25PM; Confirmation ID 168541351";
        }

        private void ListView1_ItemCheck(object sender, ItemCheckEventArgs e)
        {

        }

        private void ComboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            this.tabControl1.Visible = true;
            this.lblTitle.Text = this.comboBox1.Text + " - Open Invoices";
            this.lblTitle2.Text = this.comboBox1.Text + " - Submitted Invoices";
            getListView1(this.comboBox1.Text);
        }

        private void ListView1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            frmPop pop = new frmPop();
            pop.Show();
        }

        private void btnImportRawData_Click(object sender, EventArgs e)
        {

            if (backgroundWorker1.IsBusy != true)
            {
                // create a new instance of the alert form
                progressDlg = new frmProgress();
                progressDlg.StartPosition = FormStartPosition.CenterParent;

                // event handler for the Cancel button in AlertForm
                progressDlg.Canceled += new EventHandler<EventArgs>(buttonCancel_Click);
                progressDlg.Show(this);
                // Start the asynchronous operation.
                backgroundWorker1.RunWorkerAsync();
            }
            
        }

        // This event handler cancels the backgroundworker, fired from Cancel button in AlertForm.
        private void buttonCancel_Click(object sender, EventArgs e)
        {
            if (backgroundWorker1.WorkerSupportsCancellation == true)
            {
                // Cancel the asynchronous operation.
                backgroundWorker1.CancelAsync();
                // Close the AlertForm
                progressDlg.Close();
            }

            if (backgroundWorker2.WorkerSupportsCancellation == true)
            {
                // Cancel the asynchronous operation.
                backgroundWorker2.CancelAsync();
                // Close the AlertForm
                progressDlg.Close();
            }
        }

        // This event handler is where the time-consuming work is done.
        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            BackgroundWorker worker = sender as BackgroundWorker;

            //Read the data from the tblRawData table
            SQLiteConnection sqlite_conn;
            sqlite_conn = CreateConnection();

            SQLiteDataReader sqlite_datareader;
            SQLiteCommand sqlite_cmd;
            sqlite_cmd = sqlite_conn.CreateCommand();

            sqlite_cmd.CommandText = "SELECT a.PARTY_NAME, a.ADDRESS2, a.CITY, a.POSTAL_CODE, a.TRX_DATE, a.TRX_NUMBER, a.INVOICE_CURRENCY_CODE, a.AMOUNT_DUE_ORIGINAL, a.QUANTITY_INVOICED, a.DESCRIPTION, a.AMOUNT_DUE_ORIGINAL, a.UNIT_SELLING_PRICE, a.PURCHASE_ORDER, a.SEGMENT1, a.NAME FROM tblRawData a LEFT JOIN tblTest00 b ON a.TRX_NUMBER = b.InvoiceNumber WHERE b.InvoiceNumber IS NULL;";

            sqlite_datareader = sqlite_cmd.ExecuteReader();

            DataTable dt = new DataTable();
            dt.Load(sqlite_datareader);
            int numRows = dt.Rows.Count;
            int count = 0;

            foreach (DataRow row in dt.Rows)
            {
                string Name = Convert.ToString(row[0]);
                string Address = Convert.ToString(row[1]) + ", " + Convert.ToString(row[2]) + ", " + Convert.ToString(row[3]);
                string InvoiceDate = Convert.ToString(row[4]);
                string InvoiceNumber = Convert.ToString(row[5]);
                string Currency = Convert.ToString(row[6]);
                string NetTotal = Convert.ToString(row[7]);
                string Quantity = Convert.ToString(row[8]);
                string Description = Convert.ToString(row[9]);
                string LineNetAmount = Convert.ToString(row[7]);
                string UnitPrice = Convert.ToString(row[10]);
                string UnitOfMeasure = "ELB";
                string InvoiceType = "380";
                string PoNumber = Convert.ToString(row[11]);
                string BuyerID = "AAA791040983";
                string InvoiceGross = Convert.ToString(row[7]);
                string PartNumber = Convert.ToString(row[12]);
                string RemitToAddress = "38889 Highway 58 Buttonwillow, CA 93206";
                string ShipTo = "Conagra Foods 1023 Fourth Street Council Bluff, IA 51501";
                string OriginalInvoiceNumber = Convert.ToString(row[5]);
                string PaymentTerms = Convert.ToString(row[13]);
                string TaxCategoryCode1 = "US4";

                SQLiteCommand insertSQL = new SQLiteCommand("INSERT INTO tblTest00 (Name, BankAddress, InvoiceDate, InvoiceNumber, Currency, InvoiceNetAmount, Quantity, CarriageDescription, LineNetAmount, UnitPrice, UnitOfMeasure, InvoiceType, PONumber, BuyerID, InvoiceGrossAmount, SupplierPartNum, RemitToStreet1, ShipToAdd1, OriginalInvoiceNumber, PaymentTerms, TaxCategoryCode1) VALUES (@Name, @Address, @InvoiceDate, @InvoiceNumber, @Currency, @NetTotal, @Quantity, @Description, @LineNetAmount, @UnitPrice, @UnitOfMeasure, @InvoiceType, @PoNumber, @BuyerID, @InvoiceGross, @PartNumber, @RemitToAddress, @ShipTo, @OriginalInvoiceNumber, @PaymentTerms, @TaxCategoryCode1)", sqlite_conn);
                insertSQL.Parameters.Add(new SQLiteParameter("@Name", Name));
                insertSQL.Parameters.Add(new SQLiteParameter("@Address", Address));
                insertSQL.Parameters.Add(new SQLiteParameter("@InvoiceDate", InvoiceDate));
                insertSQL.Parameters.Add(new SQLiteParameter("@InvoiceNumber", InvoiceNumber));
                insertSQL.Parameters.Add(new SQLiteParameter("@Currency", Currency));
                insertSQL.Parameters.Add(new SQLiteParameter("@NetTotal", NetTotal));
                insertSQL.Parameters.Add(new SQLiteParameter("@Quantity", Quantity));
                insertSQL.Parameters.Add(new SQLiteParameter("@Description", Description));
                insertSQL.Parameters.Add(new SQLiteParameter("@LineNetAmount", LineNetAmount));
                insertSQL.Parameters.Add(new SQLiteParameter("@UnitPrice", UnitPrice));
                insertSQL.Parameters.Add(new SQLiteParameter("@UnitOfMeasure", UnitOfMeasure));
                insertSQL.Parameters.Add(new SQLiteParameter("@InvoiceType", InvoiceType));
                insertSQL.Parameters.Add(new SQLiteParameter("@PoNumber", PoNumber));
                insertSQL.Parameters.Add(new SQLiteParameter("@BuyerID", BuyerID));
                insertSQL.Parameters.Add(new SQLiteParameter("@InvoiceGross", InvoiceGross));
                insertSQL.Parameters.Add(new SQLiteParameter("@PartNumber", PartNumber));
                insertSQL.Parameters.Add(new SQLiteParameter("@RemitToAddress", RemitToAddress));
                insertSQL.Parameters.Add(new SQLiteParameter("@ShipTo", ShipTo));
                insertSQL.Parameters.Add(new SQLiteParameter("@OriginalInvoiceNumber", OriginalInvoiceNumber));
                insertSQL.Parameters.Add(new SQLiteParameter("@PaymentTerms", PaymentTerms));
                insertSQL.Parameters.Add(new SQLiteParameter("@TaxCategoryCode1", TaxCategoryCode1));

                try
                {
                    insertSQL.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    throw new Exception(ex.Message);
                }

                count++;
                if (worker.CancellationPending == true)
                {
                    e.Cancel = true;
                    break;
                }
                else
                {
                    // Perform a time consuming operation and report progress.
                    worker.ReportProgress(count * 100 / numRows);
                }
            }
            sqlite_conn.Close();
            
        }

        // This event handler updates the progress.
        private void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            // Show the progress in main form (GUI)
            labelResult.Text = (e.ProgressPercentage.ToString() + "%");

            // Pass the progress to AlertForm label and progressbar
            progressDlg.Message = "In progress, please wait... " + e.ProgressPercentage.ToString() + "%";
            progressDlg.ProgressValue = e.ProgressPercentage;
        }

        // This event handler deals with the results of the background operation.
        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Cancelled == true)
            {
                labelResult.Text = "Canceled!";
            }
            else if (e.Error != null)
            {
                labelResult.Text = "Error: " + e.Error.Message;
            }
            else
            {
                labelResult.Text = "Done!";
            }
            // Close the AlertForm
            progressDlg.Close();
        }

        private void backgroundWorker2_DoWork(object sender, DoWorkEventArgs e)
        {
            BackgroundWorker worker = sender as BackgroundWorker;

            SQLiteConnection sqlite_conn;
            sqlite_conn = CreateConnection();
            SQLiteDataReader sqlite_datareader;
            SQLiteCommand sqlite_cmd;
            sqlite_cmd = sqlite_conn.CreateCommand();

            sqlite_cmd.CommandText = "SELECT * FROM tblTest00";

            sqlite_datareader = sqlite_cmd.ExecuteReader();

            DataTable dt = new DataTable();
            dt.Clear();

            var columns = new List<string>();

            for (int i = 0; i < sqlite_datareader.FieldCount; i++)
            {
                dt.Columns.Add(sqlite_datareader.GetName(i));
            }

            int count = 0;

            while (sqlite_datareader.Read())
            {
                if (count < 50000)
                    count++;
                if (worker.CancellationPending == true)
                {
                    e.Cancel = true;
                    break;
                }
                else
                {
                    // Perform a time consuming operation and report progress.
                    worker.ReportProgress(count / 100);
                }

                DataRow _ravi = dt.NewRow();
                for (int i = 0; i < sqlite_datareader.FieldCount; i++)
                {
                    _ravi[sqlite_datareader.GetName(i)] = FixDecimals(Convert.ToString(sqlite_datareader.GetValue(i)));
                }
                dt.Rows.Add(_ravi);

            }

            // Write the datatable to csv file

            WriteToCsvFile(dt, csvFilePath);

            sqlite_conn.Close();
        }

        private void backgroundWorker2_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            // Show the progress in main form (GUI)
            labelResult.Text = (e.ProgressPercentage.ToString() + "%");

            // Pass the progress to AlertForm label and progressbar
            progressDlg.Message = "In progress, please wait... " + e.ProgressPercentage.ToString() + "%";
            progressDlg.ProgressValue = e.ProgressPercentage;
        }

        private void backgroundWorker2_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Cancelled == true)
            {
                labelResult.Text = "Canceled!";
            }
            else if (e.Error != null)
            {
                labelResult.Text = "Error: " + e.Error.Message;
            }
            else
            {
                labelResult.Text = "Done!";
            }
            // Close the AlertForm
            progressDlg.Close();
        }
    }
}
