using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Xml;
using System.Data.SqlClient;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Configuration;
using RestSharp;
using System.Drawing.Imaging;
using Limilabs.Barcode;
using System.Runtime.Serialization.Formatters.Binary;
using System.Drawing.Printing;
using System.Media;

namespace WindowsFormsApp1
{
    public partial class frmMain : Form
    {
        public static string user = ConfigurationManager.AppSettings["userName"];
        public static string passWord = ConfigurationManager.AppSettings["passWord"];
        public static string Token = ConfigurationManager.AppSettings["Token"];
        public static string websiteURL = null;
        //static private string websiteURL = "http://www.joinseattle.com";
        //static private string websiteURL = "http://magento-211349-651516.cloudwaysapps.com";
        //static private string websiteURL = "http://magento-211349-655102.cloudwaysapps.com/";
        public string strServerName;
        public string strDBName;
        public string connectionString;
        public int printPageNo;
        public int dataCount;
        public string strStorename;
        public string strPrinterName = null;
        public Boolean IsReading = true;

        DateTime _lastKeystroke = new DateTime(0);
        List<char> _barcode = new List<char>(10);

        public void sqlExecute(string strSQLString)
        {

            try
            {
                SqlConnection adoCon = new SqlConnection();
                adoCon.ConnectionString = connectionString;
                adoCon.Open();
                SqlCommand cmd2 = new SqlCommand();
                cmd2.Connection = adoCon;
                cmd2.CommandText = strSQLString;
                cmd2.ExecuteReader();
                adoCon.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }
        public string sqlDepartmentRead(string strSQLString)
        {

            try
            {
                SqlConnection adoCon = new SqlConnection();
                adoCon.ConnectionString = connectionString;
                adoCon.Open();
                SqlCommand command = new SqlCommand("SELECT [Description] FROM [Departments] where [Dept_ID]=" + "'" + strSQLString + "'", adoCon);
                SqlDataReader reader = command.ExecuteReader();

                if (reader.HasRows)
                {
                    while (reader.Read())
                    {
                        return reader.GetString(0);
                    }
                }

                reader.Close();
                return "No Departments";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return "error";
            }

        }

        private void InitializeDataGridView()
        {
            try
            {
                IsReading = true;

                string strSQLString;
                // Set up the DataGridView.
                {
                    var withBlock = this.DataGridView1;
                    // Automatically generate the DataGridView columns.
                    withBlock.AutoGenerateColumns = true;
                    strSQLString = " SELECT [Invoice_Number],[OnHoldID],[Name] FROM [Invoice_OnHold]";

                    BindingSource3.DataSource = GetData(strSQLString, strServerName);
                    withBlock.DataSource = BindingSource3;

                    // Automatically resize the visible rows.
                    withBlock.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.DisplayedCellsExceptHeaders;

                    // Set the DataGridView control's border.
                    withBlock.BorderStyle = BorderStyle.Fixed3D;

                    // Put the cells in edit mode when user enters them.
                    withBlock.EditMode = DataGridViewEditMode.EditOnEnter;
                }

                IsReading = false;
            }


            catch (SqlException ex)
            {
                MessageBox.Show("Comcash Data Connection Error." + ex.Errors, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        private void loadInvoiceItem(string strInvoiceNumber)
        {
            try
            {
                string strSQLString = "SELECT [LineNum],[ItemNum],[Quantity] ,[DiffItemName],[origPricePer] FROM [Invoice_Itemized] where Invoice_Number = " + strInvoiceNumber;
                SqlConnection adoCon = new SqlConnection();
                adoCon.ConnectionString = connectionString;
                adoCon.Open();
                SqlCommand command = new SqlCommand(strSQLString, adoCon);
                SqlDataReader reader = command.ExecuteReader();

                if (reader.HasRows)
                {
                    dataGridView2.Rows.Clear();
                    dataGridView2.ColumnCount = 7;

                    dataGridView2.Columns[0].Name = "LineNum";
                    dataGridView2.Columns[1].Name = "ItemNum";
                    dataGridView2.Columns[2].Name = "DiffItemName";
                    dataGridView2.Columns[3].Name = "Price";
                    dataGridView2.Columns[4].Name = "OrderQty";
                    dataGridView2.Columns[5].Name = "Count";
                    dataGridView2.Columns[6].Name = "Need More";

                    while (reader.Read())
                    {
                        string strColumn1 = reader["LineNum"].ToString().Trim();
                        string strColumn2 = reader["ItemNum"].ToString().Trim();
                        string strColumn3 = reader["DiffItemName"].ToString().Trim();
                        string strColumn4 = reader["origPricePer"].ToString().Trim();
                        string strColumn5 = reader["Quantity"].ToString().Trim();

                        dataGridView2.Rows.Add(strColumn1, strColumn2, strColumn3, decimal.Parse(strColumn4).ToString("0.00"), decimal.Parse(strColumn5).ToString("0.##"));

                    }
                }

                reader.Close();
            }

            catch (SqlException ex)
            {
                MessageBox.Show("Data Connection Error." + ex.Errors, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        public DataTable GetData(string sqlCommand, string strDBIP)
        {

            // strDBIP = "SI-PC"
            // Dim connectionString As String = "Data Source='" & strDBIP & "\SQLEXPRESS';Initial Catalog=comcash;Persist Security Info=True;User ID=ccsa;Password=ccsa"
            connectionString = "Data Source='" + strDBIP + @"\pcamerica';Initial Catalog=" + strDBName
                + ";Persist Security Info=True;User ID=sa;Password=pcAmer1ca";

            SqlConnection ComcashConnection = new SqlConnection(connectionString);

            SqlCommand command = new SqlCommand(sqlCommand, ComcashConnection);
            SqlDataAdapter adapter = new SqlDataAdapter();
            adapter.SelectCommand = command;

            DataTable table = new DataTable();
            table.Locale = System.Globalization.CultureInfo.InvariantCulture;
            adapter.Fill(table);

            return table;
        }



        private void SaveToXML(string strDBservername, string strDbName)
        {

            try
            {
                // Create an Xml document instance and load XML data.
                XmlDocument doc = new XmlDocument();
                string Path = Directory.GetCurrentDirectory() + @"\DBServer.xml";
                bool filenameExists;

                filenameExists = File.Exists(Path);
                if (filenameExists)
                    File.Delete(Path);
                if (true)
                {
                    XmlTextWriter createFile = new XmlTextWriter(Path, Encoding.Unicode);
                    createFile.WriteStartDocument(false);
                    // createFile.WriteDocType("Invoice", Nothing, Nothing, Nothing)
                    createFile.WriteComment("This file represents pcAmerica POS database");
                    createFile.WriteStartElement("pcAmerica");
                    createFile.Flush();
                    createFile.Close();

                    doc.Load(Path);
                    XmlElement newElem = doc.CreateElement("Company");

                    XmlAttribute newAttr = doc.CreateAttribute("info");
                    newAttr.Value = "pcAmerica";
                    newElem.Attributes.Append(newAttr);

                    // Create the child nodes. The following example shows various ways to add child nodes.
                    string strInnerXml;
                    strInnerXml = "<DBSERVER></DBSERVER>";
                    strInnerXml = strInnerXml + "<DBNAME></DBNAME>";
                    strInnerXml = strInnerXml + "<StoreName></StoreName>";
                    strInnerXml = strInnerXml + "<PrinterName></PrinterName>";
                    strInnerXml = strInnerXml + "<Help></Help>";

                    newElem.InnerXml = strInnerXml;

                    doc.DocumentElement.AppendChild(newElem);
                    doc.PreserveWhitespace = true;
                    XmlTextWriter wrtr = new XmlTextWriter(Path, Encoding.Unicode);
                    doc.WriteTo(wrtr);
                    // doc.WriteTo(CryptoZ.Encrypt(strInnerXml, "MyKey"))
                    wrtr.Close();
                }
                doc.Load(Path);
                XmlNodeList nodeList = doc.SelectNodes("//Company");

                foreach (XmlNode node in nodeList)
                {
                    node.SelectSingleNode("DBSERVER").InnerText = txtDBIP.Text.Trim();
                    node.SelectSingleNode("DBNAME").InnerText = txtDBname.Text.Trim();
                    node.SelectSingleNode("PrinterName").InnerText = strPrinterName;
                }
                doc.Save(Path);
                strDBservername = txtDBIP.Text.Trim();
                strDbName = txtDBname.Text.Trim();
            }
            catch (XmlException xmlex)
            {
                MessageBox.Show(xmlex.Message);
            }
            // UserLoad_From_SQL = False
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }


        public frmMain()
        {
            InitializeComponent();
            try
            {
                this.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.frmMain_KeyPress);

                //  Create an Xml document instance and load XML data.
                XmlDocument doc = new XmlDocument();
                string Path = (Directory.GetCurrentDirectory() + "\\DBServer.xml");
                if (File.Exists(Path))
                {
                    doc.Load(Path);
                    XmlNodeList nodeList = doc.SelectNodes("//Company");
                    foreach (XmlNode node in nodeList)
                    {
                        strServerName = node.SelectSingleNode("DBSERVER").InnerText;
                        txtDBIP.Text = strServerName;
                        strDBName = node.SelectSingleNode("DBNAME").InnerText;
                        txtDBname.Text = strDBName;
                        strStorename = node.SelectSingleNode("StoreName").InnerText;
                        strPrinterName = node.SelectSingleNode("PrinterName").InnerText;
                       
                    }

                    // doc.Save(Path)
                }




                connectionString = ("Data Source=\'"
                            + (strServerName + ("\\pcamerica\';Initial Catalog="
                            + (strDBName + ";Persist Security Info=True;User ID=sa;Password=pcAmer1ca"))));

            }
            catch (SqlException ex)
            {
                MessageBox.Show(("Comcash Data Connection Error." + ex.Errors), "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                //  System.Threading.Thread.CurrentThread.Abort()
            }
        }




        private void SaveDBServerToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SaveToXML(txtDBIP.Text.Trim(), txtDBname.Text.Trim());
        }




        private void btnCREConnect_Click(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;
            InitializeDataGridView();
            Cursor.Current = Cursors.Default;
        }


        private void InitializePrintPreviewDialog()
        {

            // Create a new PrintPreviewDialog using constructor.
            this.PrintPreviewDialog1 = new PrintPreviewDialog();

            //Set the size, location, and name.
            this.PrintPreviewDialog1.ClientSize =
                new System.Drawing.Size(400, 300);
            this.PrintPreviewDialog1.Location =
                new System.Drawing.Point(29, 29);
            this.PrintPreviewDialog1.Name = "PrintPreviewDialog1";

            // Associate the event-handling method with the 
            // document's PrintPage event.
            //this.PrintDocument1.PrintPage += new System.Drawing.Printing.PrintPageEventHandler(PrintDocument1_PrintPage);

            // Set the minimum size the dialog can be resized to.
            this.PrintPreviewDialog1.MinimumSize =
                new System.Drawing.Size(600, 300);

            // Set the UseAntiAlias property to true, which will allow the 
            // operating system to smooth fonts.
            this.PrintPreviewDialog1.UseAntiAlias = true;
        }

        private void button1_Click_1(object sender, EventArgs e)
        {


            InitializePrintPreviewDialog();
            PrintDocument1.DefaultPageSettings.Landscape = false;
            PrintDocument1.DefaultPageSettings.PaperSize = new System.Drawing.Printing.PaperSize("Hibro", 120, 85);
            PrintPreviewDialog1.Document = PrintDocument1;
            PrintPreviewDialog1.ShowDialog();

        }

 
      

        private void SaveDBServerToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            SaveToXML(txtDBIP.Text.Trim(), txtDBname.Text.Trim());
        }

 

        private void button2_Click(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;
            InitializeDataGridView();
            Cursor.Current = Cursors.Default;
        }

        private void DataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            if (IsReading == false && DataGridView1.CurrentCell != null && DataGridView1.RowCount > DataGridView1.CurrentCell.RowIndex + 1)
            {
                int indexrow = DataGridView1.CurrentCell.RowIndex;
                string strSKU = DataGridView1.Rows[indexrow].Cells[0].Value.ToString();
                txtInvoiceNumber.Text = strSKU;
                loadInvoiceItem(strSKU);
            }


        }


        private void txtItemNumber_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)Keys.Return && dataGridView2.Rows.Count >= 2)
            {
                String searchValue = txtItemNumber.Text;
                Boolean isFinded = false;

                int i = -1;
                for (i = 0; (i <= (dataGridView2.Rows.Count - 2)); i++)
                {
                    if (dataGridView2.Rows[i].Cells[1].Value.ToString().Equals(searchValue))
                    {
                        dataGridView2.Rows[i].Cells[1].Style.BackColor = Color.LightBlue;

                        dataGridView2.Rows[i].Selected = true;

                        dataGridView2.Rows[i].Cells[5].Value =Convert.ToInt16(dataGridView2.Rows[i].Cells[5].Value) + 1;
                        dataGridView2.Rows[i].Cells[6].Value = Convert.ToInt16(dataGridView2.Rows[i].Cells[4].Value) - Convert.ToInt16(dataGridView2.Rows[i].Cells[5].Value);

                        if (Convert.ToInt16(dataGridView2.Rows[i].Cells[6].Value) == 0)
                        {
                            dataGridView2.Rows[i].Cells[4].Style.BackColor = Color.LightGreen;
                            dataGridView2.Rows[i].Cells[5].Style.BackColor = Color.LightGreen;
                            dataGridView2.Rows[i].Cells[6].Style.BackColor = Color.LightGreen;
                        }
                        if (Convert.ToInt16(dataGridView2.Rows[i].Cells[6].Value) < 0)
                        {
                            SystemSounds.Hand.Play();
                            MessageBox.Show("Too Much already full", "Check", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                            dataGridView2.Rows[i].Cells[4].Style.BackColor = Color.Red;
                            dataGridView2.Rows[i].Cells[5].Style.BackColor = Color.Red;
                            dataGridView2.Rows[i].Cells[6].Style.BackColor = Color.Red;
                            dataGridView2.Rows[i].Cells[6].Value = 0;
                        }
                        isFinded = true;
                        break;
                    }
                }
                if (isFinded == false)
                {
                    dataGridView2.Rows[i].Selected = true;
                    MessageBox.Show("Item Not Ordered", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }
                txtItemNumber.SelectAll();
                txtItemNumber.Focus();
                e.Handled = true;
            }
            
        }

        private void txtInvoiceNumber_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)Keys.Return && txtInvoiceNumber.Text != "")
            {
                loadInvoiceItem(txtInvoiceNumber.Text);
                e.Handled = true;
            }
        }

        private void txtItemNumber_Leave(object sender, EventArgs e)
        {
            //txtItemNumber.Text = "";
        }

      
        private void frmMain_KeyPress(object sender, KeyPressEventArgs e)
        {
            // check timing (keystrokes within 100 ms)
            TimeSpan elapsed = (DateTime.Now - _lastKeystroke);
            if (elapsed.TotalMilliseconds > 100)
                _barcode.Clear();

            // record keystroke & timestamp
            _barcode.Add(e.KeyChar);
            _lastKeystroke = DateTime.Now;

            // process barcode
            if (e.KeyChar == 13 && _barcode.Count > 0)
            {
                string msg = new String(_barcode.ToArray());
                txtItemNumber.Text = msg;
                _barcode.Clear();
            }
        }

        private void dataGridView2_KeyPress(object sender, KeyPressEventArgs e)
        {
            txtItemNumber.Text = txtItemNumber.Text + e.KeyChar;
        }
    }
}
