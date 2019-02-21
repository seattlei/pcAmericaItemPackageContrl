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
using System.Runtime.Serialization.Formatters.Binary;
using System.Drawing.Printing;
using System.Media;
using System.Xml.Serialization;

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

        DataSet ds = new DataSet();
        DataTable dt = new DataTable();

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

                    bindingSource1.DataSource = GetData(strSQLString, strServerName);
                    withBlock.DataSource = bindingSource1;

                    // Automatically resize the visible rows.
                    withBlock.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.DisplayedCellsExceptHeaders;

                    // Set the DataGridView control's border.
                    withBlock.BorderStyle = BorderStyle.Fixed3D;

                    // Put the cells in edit mode when user enters them.
                    withBlock.EditMode = DataGridViewEditMode.EditOnEnter;
                    DataGridView1.Columns[0].Width = 80;
                    DataGridView1.Columns[1].Width = 180;
                    DataGridView1.Columns[2].Width = 180;
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

                XmlDocument doc = new XmlDocument();
                string Path = Directory.GetCurrentDirectory() + @"\" + strInvoiceNumber + ".xml";
                bool filenameExists;

                filenameExists = File.Exists(Path);
                

                if (reader.HasRows)
                {
                    if (filenameExists)
                    {
                        dataGridView2.DataSource = null;
                        dataGridView2.Columns.Clear();
                        dataGridView2.Rows.Clear();
                        ds.Reset();
                        ds.ReadXml(Path);
                        dataGridView2.DataSource = ds.Tables[0];

                        foreach (DataGridViewRow row in dataGridView2.Rows)
                        {
                            row.Cells[6].Style.BackColor = Color.Red;
                        }
                        try
                        {
                            dataGridView2.Columns[0].Width = 50;
                            dataGridView2.Columns[1].Width = 180;
                            dataGridView2.Columns[2].Width = 80;
                            dataGridView2.Columns[3].Width = 60;
                            dataGridView2.Columns[4].Width = 60;
                            dataGridView2.Columns[5].Width = 60;
                            dataGridView2.Columns[6].Width = 400;
                            dataGridView2.Columns[2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                            dataGridView2.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                            dataGridView2.Columns[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                            dataGridView2.Columns[5].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                            if (reader.HasRows)
                            {
                                while (reader.Read())
                                {
                                    int rowIndex = -1;
                                    string strColumn1 = reader["LineNum"].ToString().Trim();
                                    string strColumn2 = reader["ItemNum"].ToString().Trim();
                                    string strColumn3 = reader["origPricePer"].ToString().Trim();
                                    string strColumn4 = reader["Quantity"].ToString().Trim();
                                    string strColumn5 = reader["DiffItemName"].ToString().Trim();

                                    dataGridView2.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
                                    foreach (DataGridViewRow row in dataGridView2.Rows)
                                    {   
                                        if (row.Cells[1].Value.ToString().Equals(strColumn2))
                                        {
                                            rowIndex = row.Index;
                                            row.Selected = true;
                                            break;
                                        }
                                    }
                                    if (rowIndex == -1)
                                    {
                                        dataGridView2.Rows.Add(strColumn1, strColumn2, decimal.Parse(strColumn3).ToString("0.00"), decimal.Parse(strColumn4).ToString("0.##"), 0, decimal.Parse(strColumn4).ToString("0.##"), strColumn5);
                                        //dataGridView2.Rows[dataGridView2.Rows.Count - 1].Selected = true;
                                        dataGridView2.Rows[dataGridView2.Rows.Count - 1].Cells[6].Style.BackColor = Color.CadetBlue;
                                    }
                                    else
                                    {
                                        dataGridView2.Rows[rowIndex].Cells[6].Style.BackColor = Color.White;
                                        dataGridView2.Rows[rowIndex].Cells[3].Value = decimal.Parse(strColumn4).ToString("0.##");
                                        Single CountNeed = Convert.ToSingle(strColumn4) - Convert.ToSingle(dataGridView2.Rows[rowIndex].Cells[4].Value);
                                        dataGridView2.Rows[rowIndex].Cells[5].Value = CountNeed;
                                        if (CountNeed < 0)
                                        {
                                            dataGridView2.Rows[rowIndex].Cells[3].Style.BackColor = Color.Red;
                                            dataGridView2.Rows[rowIndex].Cells[4].Style.BackColor = Color.Red;
                                            dataGridView2.Rows[rowIndex].Cells[5].Style.BackColor = Color.Red;
                                        }
                                    }

                                }
                            }

                            reader.Close();
                        }
                        catch (SqlException ex)
                        {
                            MessageBox.Show("Data Connection Error." + ex.Errors, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        }

                    }

                    else
                    {
                        dataGridView2.DataSource = null;
                        dataGridView2.BackgroundColor = Color.White;
                        dataGridView2.Rows.Clear();
                        dataGridView2.ColumnCount = 7;

                        dataGridView2.Columns[0].Width = 50;
                        dataGridView2.Columns[1].Width = 180;
                        dataGridView2.Columns[2].Width = 80;
                        dataGridView2.Columns[3].Width = 60;
                        dataGridView2.Columns[4].Width = 60;
                        dataGridView2.Columns[5].Width = 60;
                        dataGridView2.Columns[6].Width = 400;

                        dataGridView2.Columns[0].Name = "LineNum";
                        dataGridView2.Columns[1].Name = "ItemNum";
                        dataGridView2.Columns[2].Name = "Price";
                        dataGridView2.Columns[3].Name = "OrderQty";
                        dataGridView2.Columns[4].Name = "Count";
                        dataGridView2.Columns[5].Name = "NeedMore";
                        dataGridView2.Columns[6].Name = "DiffItemName";
                        dataGridView2.Columns[2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                        dataGridView2.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                        dataGridView2.Columns[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                        dataGridView2.Columns[5].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;


                        while (reader.Read())
                        {
                            string strColumn1 = reader["LineNum"].ToString().Trim();
                            string strColumn2 = reader["ItemNum"].ToString().Trim();
                            string strColumn3 = reader["origPricePer"].ToString().Trim();
                            string strColumn4 = reader["Quantity"].ToString().Trim();
                            string strColumn6 = reader["DiffItemName"].ToString().Trim();

                            dataGridView2.Rows.Add(strColumn1, strColumn2, decimal.Parse(strColumn3).ToString("0.00"), decimal.Parse(strColumn4).ToString("0.##"), 0, decimal.Parse(strColumn4).ToString("0.##"), strColumn6);

                        }
                    }

                    
                    
                }
                //dataGridView2.DataSource = gridTable;
                
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
                var version = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
                this.Text = String.Format("pcAmerica Pakage Double Checker  Version {0}", version);

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
                    DataGridView1.DefaultCellStyle.Font= new Font("Arial", 20F, GraphicsUnit.Pixel);
                    dataGridView2.DefaultCellStyle.Font = new Font("Arial", 20F, GraphicsUnit.Pixel);
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
            try
            {

                if (IsReading == false && DataGridView1.CurrentCell != null && DataGridView1.RowCount > DataGridView1.CurrentCell.RowIndex)
                {
                    int indexrow = DataGridView1.CurrentCell.RowIndex;
                    string strSKU = DataGridView1.Rows[indexrow].Cells[0].Value.ToString();
                    txtInvoiceNumber.Text = strSKU;
                    txtItemNumber.Text = "";
                    loadInvoiceItem(strSKU);
                }
                txtItemNumber.Focus();
                txtItemNumber.SelectAll();

            }
            catch
            {

            }
        }


        private void txtItemNumber_KeyPress(object sender, KeyPressEventArgs e)
        {
            try
            {
                if (e.KeyChar == (char)Keys.Return && dataGridView2.Rows.Count >= 1 && txtItemNumber.Text != "")
                {
                    String searchValue = txtItemNumber.Text;
                    string searchValue2 = null;

                    string strSQLString = "SELECT s.ItemNum FROM Inventory_SKUs s INNER JOIN Inventory i ON s.Store_ID = i.Store_ID AND s.ItemNum = i.ItemNum WHERE s.Store_ID = '1001' AND s.ALTSku = '" + searchValue + "'";
                    SqlConnection adoCon = new SqlConnection();
                    adoCon.ConnectionString = connectionString;
                    adoCon.Open();
                    SqlCommand command = new SqlCommand(strSQLString, adoCon);
                    SqlDataReader reader = command.ExecuteReader();

                    if (reader.HasRows)
                    {
                        reader.Read();
                        searchValue2 = reader["ItemNum"].ToString().Trim();
                    }

                    Boolean isFinded = false;

                    int i = -1;
                    for (i = 0; (i <= (dataGridView2.Rows.Count - 1)); i++)
                    {
                        if (dataGridView2.Rows[i].Cells[1].Value.ToString().Equals(searchValue) || dataGridView2.Rows[i].Cells[1].Value.ToString().Equals(searchValue2))
                        {
                            dataGridView2.FirstDisplayedScrollingRowIndex = i;
                            dataGridView2.Rows[i].Cells[1].Style.BackColor = Color.LightBlue;

                            dataGridView2.Rows[i].Selected = true;

                            dataGridView2.Rows[i].Cells[4].Value = Convert.ToInt16(dataGridView2.Rows[i].Cells[4].Value) + 1;
                            dataGridView2.Rows[i].Cells[5].Value = Convert.ToInt16(dataGridView2.Rows[i].Cells[3].Value) - Convert.ToInt16(dataGridView2.Rows[i].Cells[4].Value);

                            if (Convert.ToInt16(dataGridView2.Rows[i].Cells[5].Value) == 0)
                            {
                                dataGridView2.Rows[i].Cells[3].Style.BackColor = Color.LightGreen;
                                dataGridView2.Rows[i].Cells[4].Style.BackColor = Color.LightGreen;
                                dataGridView2.Rows[i].Cells[5].Style.BackColor = Color.LightGreen;
                            }
                            if (Convert.ToInt16(dataGridView2.Rows[i].Cells[5].Value) < 0)
                            {
                                SystemSounds.Hand.Play();
                                MessageBox.Show("Too Much already full", "Check", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                                dataGridView2.Rows[i].Cells[3].Style.BackColor = Color.Red;
                                dataGridView2.Rows[i].Cells[4].Style.BackColor = Color.Red;
                                dataGridView2.Rows[i].Cells[5].Style.BackColor = Color.Red;
                                //dataGridView2.Rows[i].Cells[5].Value = 0;
                            }
                            isFinded = true;
                            break;
                        }
                    }
                    if (isFinded == false)
                    {
                        dataGridView2.Rows[i-1].Selected = true;
                        MessageBox.Show("Item Not Ordered", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    }
                    else
                    {
                        SaveInvoiceToXML2(txtInvoiceNumber.Text.ToString());
                    }
                    e.Handled = true;
                    txtItemNumber.Text = "";
                    txtItemNumber.SelectAll();
                    txtItemNumber.Focus();
                }
                
                
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);

            }

        }
       

        private void SaveInvoiceToXML(string InvoiceNumber)
        {

            try
            {
                // Create an Xml document instance and load XML data.
                XmlDocument doc = new XmlDocument();
                string Path = Directory.GetCurrentDirectory() + @"\" + InvoiceNumber + ".xml";
                bool filenameExists;

                filenameExists = File.Exists(Path);
                if (filenameExists)
                    File.Delete(Path);

                XmlTextWriter createFile = new XmlTextWriter(Path, Encoding.Unicode);
                createFile.WriteStartDocument();
                // createFile.WriteDocType("Invoice", Nothing, Nothing, Nothing)
                createFile.WriteComment("This file represents pcAmerica POS data");
                
                createFile.WriteStartElement("Table");

                int i = -1;
                for (i = 0; (i <= (dataGridView2.Rows.Count - 2)); i++)
                {
                    createFile.WriteStartElement("Product");
                    createFile.WriteStartElement("ItemNum");
                    createFile.WriteString(Convert.ToString(dataGridView2.Rows[i].Cells[1].Value));
                    createFile.WriteEndElement();
                    createFile.WriteStartElement("IteOrderQtymNum");
                    createFile.WriteString(Convert.ToString(dataGridView2.Rows[i].Cells[3].Value));
                    createFile.WriteEndElement();
                    createFile.WriteStartElement("PickupCount");
                    createFile.WriteString(Convert.ToString(dataGridView2.Rows[i].Cells[4].Value));
                    createFile.WriteEndElement();
                    createFile.WriteEndElement();

                }
                createFile.WriteEndElement();
                createFile.WriteEndDocument();
                createFile.Flush();
                createFile.Close();

                
              
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

        private void SaveInvoiceToXML2(string InvoiceNumber)
        {

            try
            {
                XmlDocument doc = new XmlDocument();
                string Path = Directory.GetCurrentDirectory() + @"\" + InvoiceNumber + ".xml";
                bool filenameExists;

                filenameExists = File.Exists(Path);
                if (filenameExists) File.Delete(Path);
                dt.Reset();
                dt.TableName = "Double Check";
               
                foreach (DataGridViewColumn col in dataGridView2.Columns)
                {
                    dt.Columns.Add(col.Name);
                }

                foreach (DataGridViewRow row in dataGridView2.Rows)
                {
                    DataRow dRow = dt.NewRow();
                    foreach (DataGridViewCell cell in row.Cells)
                    {
                        dRow[cell.ColumnIndex] = cell.Value;
                    }
                    dt.Rows.Add(dRow);
                }

                dt.WriteXml(Path);

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
        private void txtInvoiceNumber_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)Keys.Return && txtInvoiceNumber.Text != "")
            {
                txtItemNumber.Text = "";
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

        private void btmRefresh_Click(object sender, EventArgs e)
        {

            try
            {
                string strSQLString = "SELECT [LineNum],[ItemNum],[Quantity] ,[DiffItemName],[origPricePer] FROM [Invoice_Itemized] where Invoice_Number = " + txtInvoiceNumber.Text.ToString();
                SqlConnection adoCon = new SqlConnection();
                adoCon.ConnectionString = connectionString;
                adoCon.Open();
                SqlCommand command = new SqlCommand(strSQLString, adoCon);
                SqlDataReader reader = command.ExecuteReader();

                foreach (DataGridViewRow row in dataGridView2.Rows)
                {
                    row.Cells[6].Style.BackColor = Color.Red;
                }
                
                if (reader.HasRows)
                {
                    

                    while (reader.Read())
                    {
                        int rowIndex = -1;
                        string strColumn1 = reader["LineNum"].ToString().Trim();
                        string strColumn2 = reader["ItemNum"].ToString().Trim();
                        string strColumn3 = reader["origPricePer"].ToString().Trim();
                        string strColumn4 = reader["Quantity"].ToString().Trim();
                        string strColumn5 = reader["DiffItemName"].ToString().Trim();

                        dataGridView2.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
                        foreach (DataGridViewRow row in dataGridView2.Rows)
                        {
                            if (row.Cells[1].Value.ToString().Equals(strColumn2))
                            {
                                rowIndex = row.Index;
                                row.Selected = true;
                                break;
                            }
                        }
                        if (rowIndex == -1)
                        {
                            dataGridView2.Rows.Add(strColumn1, strColumn2, decimal.Parse(strColumn3).ToString("0.00"), decimal.Parse(strColumn4).ToString("0.##"), 0, decimal.Parse(strColumn4).ToString("0.##"), strColumn5);
                            //dataGridView2.Rows[dataGridView2.Rows.Count - 1].Selected = true;
                            dataGridView2.Rows[dataGridView2.Rows.Count - 1].Cells[6].Style.BackColor = Color.CadetBlue;
                        }
                        else
                        {
                            dataGridView2.Rows[rowIndex].Cells[6].Style.BackColor = Color.White;
                            dataGridView2.Rows[rowIndex].Cells[3].Value = decimal.Parse(strColumn4).ToString("0.##");
                            Single CountNeed = Convert.ToSingle( strColumn4) - Convert.ToSingle(dataGridView2.Rows[rowIndex].Cells[4].Value);
                            dataGridView2.Rows[rowIndex].Cells[5].Value = CountNeed;
                        }

                    }
                }

                reader.Close();
            }

            catch (SqlException ex)
            {
                MessageBox.Show("Data Connection Error." + ex.Errors, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }


        private void dataGridView2_SelectionChanged(object sender, EventArgs e)
        {
            txtItemNumber.SelectAll();
            txtItemNumber.Focus();
        }

        private void txtItemNumber_Click(object sender, EventArgs e)
        {
            txtItemNumber.Text = "";
        }

        private void button1_Click(object sender, EventArgs e)
        {
            XmlDocument doc = new XmlDocument();
            string Path = Directory.GetCurrentDirectory() + @"\" + txtInvoiceNumber.Text + ".xml";
            bool filenameExists;

            filenameExists = File.Exists(Path);
            if (filenameExists) File.Delete(Path);
            dt.Reset();
            ds.Reset();
        }

        private void dataGridView2_Click(object sender, EventArgs e)
        {
            txtItemNumber.SelectAll();
            txtItemNumber.Focus();
        }

        private void DataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            txtItemNumber.SelectAll();
            txtItemNumber.Focus();
        }

        private void frmMain_Load(object sender, EventArgs e)
        {
            txtItemNumber.SelectAll();
            txtItemNumber.Focus();
        }

        private void DataGridView1_Click(object sender, EventArgs e)
        {
            txtItemNumber.SelectAll();
            txtItemNumber.Focus();
        }
    }
}
