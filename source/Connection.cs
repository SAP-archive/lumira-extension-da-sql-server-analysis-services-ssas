using System;
using System.Drawing;
using System.Collections;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Globalization;

namespace MDXQueryBuilder
{
    /// <summary>
    /// Summary description for Connection.
    /// </summary>
    public class frmConnection : System.Windows.Forms.Form
    {
        enum Mode
        {
            PREVIEW, REFRESH, EDIT
        };
        Mode mode = Mode.PREVIEW;
        int size = 999999999;
        int iRow = -1;
        int nbRows = 0;
        string parameters = "";
        string locale = "";
        string strServer = "";
        string strUser = Environment.MachineName;
        string strPassword = "";
        string strCatalogName = "";
        string saveCatalogName = "";
        string strQuery = "";
        string connectionString = "";
        bool isCubeBuilt = false;
        string errorMessage = "";
        char cArraySplit = (char)255;
        mdxUtil util = new mdxUtil();
        formQuery _myQuery = new formQuery();
        ArrayList listCatalogs = new ArrayList();

        [DllImport("kernel32.dll")]
        public static extern bool AllocConsole();

        [DllImport("kernel32.dll")]
        public static extern bool FreeConsole();

        private System.Windows.Forms.Panel panelConnection;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox textServer;
        private System.Windows.Forms.Splitter splitter1;
        private System.Windows.Forms.TextBox textPassword;
        private System.Windows.Forms.TextBox textUser;

        private System.Windows.Forms.Button butLogin;
        private Panel panelMDX;
        public TextBox textMDXQuery;
        private Button butRunQuery;
        private Label label4;
        private Button butViewQuery;
        private ListBox listBoxCubes;

        // "localhost" "" "" "Adventure Works DW" "SELECT NON EMPTY {[Measures].[Order Quantity],[Measures].[Extended Amount],[Measures].[Total Product Cost],[Measures].[Sales Amount],[Measures].[Gross Profit],[Measures].[Freight Cost]} ON COLUMNS, NON EMPTY Crossjoin(Crossjoin(Crossjoin([Date].[Calendar Year].[Calendar Year].Members, [Date].[Calendar Quarter of Year].[Calendar Quarter of Year].Members), [Product].[Category].[Category].Members), [Product].[Subcategory].[Subcategory].Members) DIMENSION PROPERTIES PARENT_UNIQUE_NAME ON ROWS FROM [Adventure Works]"

        public frmConnection(string[] args)
        {
            //
            // Required for Windows Form Designer support
            //
            InitializeComponent();
            parseArguments(args);
        }

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        protected override void Dispose(bool disposing)
        {
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code
        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmConnection));
            this.panelConnection = new System.Windows.Forms.Panel();
            this.listBoxCubes = new System.Windows.Forms.ListBox();
            this.butViewQuery = new System.Windows.Forms.Button();
            this.butRunQuery = new System.Windows.Forms.Button();
            this.butLogin = new System.Windows.Forms.Button();
            this.label4 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.textPassword = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.textUser = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.textServer = new System.Windows.Forms.TextBox();
            this.splitter1 = new System.Windows.Forms.Splitter();
            this.panelMDX = new System.Windows.Forms.Panel();
            this.textMDXQuery = new System.Windows.Forms.TextBox();
            this.panelConnection.SuspendLayout();
            this.panelMDX.SuspendLayout();
            this.SuspendLayout();
            // 
            // panelConnection
            // 
            this.panelConnection.BackColor = System.Drawing.Color.Gray;
            this.panelConnection.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.panelConnection.Controls.Add(this.listBoxCubes);
            this.panelConnection.Controls.Add(this.butViewQuery);
            this.panelConnection.Controls.Add(this.butRunQuery);
            this.panelConnection.Controls.Add(this.butLogin);
            this.panelConnection.Controls.Add(this.label4);
            this.panelConnection.Controls.Add(this.label3);
            this.panelConnection.Controls.Add(this.textPassword);
            this.panelConnection.Controls.Add(this.label2);
            this.panelConnection.Controls.Add(this.textUser);
            this.panelConnection.Controls.Add(this.label1);
            this.panelConnection.Controls.Add(this.textServer);
            this.panelConnection.Dock = System.Windows.Forms.DockStyle.Top;
            this.panelConnection.ForeColor = System.Drawing.Color.Black;
            this.panelConnection.Location = new System.Drawing.Point(0, 0);
            this.panelConnection.Name = "panelConnection";
            this.panelConnection.Size = new System.Drawing.Size(939, 231);
            this.panelConnection.TabIndex = 7;
            // 
            // listBoxCubes
            // 
            this.listBoxCubes.BackColor = System.Drawing.Color.Gray;
            this.listBoxCubes.ForeColor = System.Drawing.Color.White;
            this.listBoxCubes.FormattingEnabled = true;
            this.listBoxCubes.ItemHeight = 15;
            this.listBoxCubes.Location = new System.Drawing.Point(517, 9);
            this.listBoxCubes.Name = "listBoxCubes";
            this.listBoxCubes.Size = new System.Drawing.Size(408, 214);
            this.listBoxCubes.TabIndex = 16;
            // 
            // butViewQuery
            // 
            this.butViewQuery.BackColor = System.Drawing.Color.Gray;
            this.butViewQuery.Enabled = false;
            this.butViewQuery.Font = new System.Drawing.Font("Arial", 9.75F);
            this.butViewQuery.ForeColor = System.Drawing.Color.White;
            this.butViewQuery.Location = new System.Drawing.Point(407, 136);
            this.butViewQuery.Name = "butViewQuery";
            this.butViewQuery.Size = new System.Drawing.Size(104, 41);
            this.butViewQuery.TabIndex = 9;
            this.butViewQuery.Text = "View query results";
            this.butViewQuery.UseVisualStyleBackColor = false;
            this.butViewQuery.Click += new System.EventHandler(this.butViewQuery_Click);
            // 
            // butRunQuery
            // 
            this.butRunQuery.BackColor = System.Drawing.Color.Gray;
            this.butRunQuery.Enabled = false;
            this.butRunQuery.Font = new System.Drawing.Font("Arial", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.butRunQuery.ForeColor = System.Drawing.Color.White;
            this.butRunQuery.Location = new System.Drawing.Point(407, 183);
            this.butRunQuery.Name = "butRunQuery";
            this.butRunQuery.Size = new System.Drawing.Size(104, 41);
            this.butRunQuery.TabIndex = 9;
            this.butRunQuery.Text = "Run Query";
            this.butRunQuery.UseVisualStyleBackColor = false;
            this.butRunQuery.Click += new System.EventHandler(this.butRunQuery_Click);
            // 
            // butLogin
            // 
            this.butLogin.BackColor = System.Drawing.Color.Gray;
            this.butLogin.Font = new System.Drawing.Font("Arial", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.butLogin.ForeColor = System.Drawing.Color.White;
            this.butLogin.Location = new System.Drawing.Point(407, 10);
            this.butLogin.Name = "butLogin";
            this.butLogin.Size = new System.Drawing.Size(104, 40);
            this.butLogin.TabIndex = 15;
            this.butLogin.Text = "&Login";
            this.butLogin.UseVisualStyleBackColor = false;
            this.butLogin.Click += new System.EventHandler(this.butLogin_Click);
            // 
            // label4
            // 
            this.label4.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)));
            this.label4.BackColor = System.Drawing.Color.Gray;
            this.label4.Font = new System.Drawing.Font("Arial", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.ForeColor = System.Drawing.Color.White;
            this.label4.Location = new System.Drawing.Point(10, 204);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(110, 20);
            this.label4.TabIndex = 12;
            this.label4.Text = "MDX QUERY";
            // 
            // label3
            // 
            this.label3.BackColor = System.Drawing.Color.Gray;
            this.label3.Font = new System.Drawing.Font("Arial", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.ForeColor = System.Drawing.Color.White;
            this.label3.Location = new System.Drawing.Point(10, 71);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(64, 24);
            this.label3.TabIndex = 12;
            this.label3.Text = "Password";
            // 
            // textPassword
            // 
            this.textPassword.BackColor = System.Drawing.Color.White;
            this.textPassword.Font = new System.Drawing.Font("Arial", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textPassword.ForeColor = System.Drawing.Color.Black;
            this.textPassword.Location = new System.Drawing.Point(74, 71);
            this.textPassword.Name = "textPassword";
            this.textPassword.PasswordChar = '*';
            this.textPassword.Size = new System.Drawing.Size(318, 21);
            this.textPassword.TabIndex = 11;
            // 
            // label2
            // 
            this.label2.BackColor = System.Drawing.Color.Gray;
            this.label2.Font = new System.Drawing.Font("Arial", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.ForeColor = System.Drawing.Color.White;
            this.label2.Location = new System.Drawing.Point(10, 40);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(64, 24);
            this.label2.TabIndex = 10;
            this.label2.Text = "User";
            // 
            // textUser
            // 
            this.textUser.BackColor = System.Drawing.Color.White;
            this.textUser.Font = new System.Drawing.Font("Arial", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textUser.ForeColor = System.Drawing.Color.Black;
            this.textUser.Location = new System.Drawing.Point(74, 40);
            this.textUser.Name = "textUser";
            this.textUser.Size = new System.Drawing.Size(318, 21);
            this.textUser.TabIndex = 9;
            // 
            // label1
            // 
            this.label1.BackColor = System.Drawing.Color.Gray;
            this.label1.Font = new System.Drawing.Font("Arial", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.ForeColor = System.Drawing.Color.White;
            this.label1.Location = new System.Drawing.Point(10, 10);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(64, 24);
            this.label1.TabIndex = 8;
            this.label1.Text = "Server";
            // 
            // textServer
            // 
            this.textServer.BackColor = System.Drawing.Color.White;
            this.textServer.Font = new System.Drawing.Font("Arial", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textServer.ForeColor = System.Drawing.Color.Black;
            this.textServer.Location = new System.Drawing.Point(74, 10);
            this.textServer.Name = "textServer";
            this.textServer.Size = new System.Drawing.Size(318, 21);
            this.textServer.TabIndex = 7;
            // 
            // splitter1
            // 
            this.splitter1.BackColor = System.Drawing.Color.Gray;
            this.splitter1.Dock = System.Windows.Forms.DockStyle.Top;
            this.splitter1.Location = new System.Drawing.Point(0, 231);
            this.splitter1.Name = "splitter1";
            this.splitter1.Size = new System.Drawing.Size(939, 4);
            this.splitter1.TabIndex = 8;
            this.splitter1.TabStop = false;
            // 
            // panelMDX
            // 
            this.panelMDX.BackColor = System.Drawing.Color.Gray;
            this.panelMDX.Controls.Add(this.textMDXQuery);
            this.panelMDX.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panelMDX.Location = new System.Drawing.Point(0, 235);
            this.panelMDX.Name = "panelMDX";
            this.panelMDX.Size = new System.Drawing.Size(939, 514);
            this.panelMDX.TabIndex = 17;
            // 
            // textMDXQuery
            // 
            this.textMDXQuery.BackColor = System.Drawing.Color.LightGray;
            this.textMDXQuery.Dock = System.Windows.Forms.DockStyle.Fill;
            this.textMDXQuery.Enabled = false;
            this.textMDXQuery.Location = new System.Drawing.Point(0, 0);
            this.textMDXQuery.Multiline = true;
            this.textMDXQuery.Name = "textMDXQuery";
            this.textMDXQuery.Size = new System.Drawing.Size(939, 514);
            this.textMDXQuery.TabIndex = 19;
            // 
            // frmConnection
            // 
            this.AutoScaleBaseSize = new System.Drawing.Size(6, 14);
            this.ClientSize = new System.Drawing.Size(939, 749);
            this.Controls.Add(this.panelMDX);
            this.Controls.Add(this.splitter1);
            this.Controls.Add(this.panelConnection);
            this.Font = new System.Drawing.Font("Arial", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(0)))), ((int)(((byte)(192)))));
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "frmConnection";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "SSAS Query";
            this.Load += new System.EventHandler(this.frmConnection_Load);
            this.panelConnection.ResumeLayout(false);
            this.panelConnection.PerformLayout();
            this.panelMDX.ResumeLayout(false);
            this.panelMDX.PerformLayout();
            this.ResumeLayout(false);

        }
        #endregion

        [STAThread]
        static void Main(string[] args)
        {
            try
            {
                Application.Run(new frmConnection(args));
            }
            catch { }

        }

        private void frmConnection_Load(object sender, EventArgs e)
        {
            textServer.Text = Environment.MachineName;
            //textUser.Text = Environment.UserName;
            util.iArrayFetch = Convert.ToInt32(999999999);
        }

        private void parseArguments(String[] args)
        {
            string strMessage = "";
            mode = Mode.PREVIEW;
            for (int i = 0; i < args.Length; i++)
            {
                strMessage += "\n" + args[i];
                if (args[i].ToLower() == "-mode" && i + 1 < args.Length)
                {
                    if (args[i + 1].ToLower() == "preview")
                        mode = Mode.PREVIEW;
                    else if (args[i + 1].ToLower() == "edit")
                        mode = Mode.EDIT;
                    else if (args[i + 1].ToLower() == "refresh")
                        mode = Mode.REFRESH;
                }
                else if (args[i].ToLower() == "-size" && i + 1 < args.Length)
                    size = Int32.Parse(args[i + 1]);
                else if (args[i].ToLower() == "-locale" && i + 1 < args.Length)
                    locale = args[i + 1];
                else if (args[i].ToLower() == "-params" && i + 1 < args.Length)
                    parameters = args[i + 1];
            }
            if (mode == Mode.EDIT || mode == Mode.REFRESH)
            {
                String[] lines = parameters.Split(';');
                for (int index = 0; index < lines.Length; index++)
                {
                    String[] tokens = lines[index].Split('=');
                    if (tokens[0].ToLower() == "server")
                        strServer = tokens[1].Replace("|~~|", "\"");
                    if (tokens[0].ToLower() == "user")
                        strUser = tokens[1].Replace("|~~|", "\"");
                    if (tokens[0].ToLower() == "password")
                        strPassword = tokens[1].Replace("|~~|", "\"");
                    if (tokens[0].ToLower() == "catalog")
                        strCatalogName = tokens[1].Replace("|~~|", "\"");
                    if (tokens[0].ToLower() == "mdx query")
                        strQuery = tokens[1].Replace("|~~|", "\"");
                }
                if (mode == Mode.REFRESH)
                {
                    connectionString = "Provider=MSOLAP.5;Data Source=" + strServer + ";Location=" + strServer + ";MDX Compatibility=2";
                    writeConsole();
                    this.Close();
                }
                else
                {
                    textServer.Text = strServer;
                    textUser.Text = strUser;
                    textPassword.Text = strPassword;
                    textMDXQuery.Text = strQuery;
                    butLogin_Click(null, null);
                }
            }
            //MessageBox.Show(strMessage);
        }

        #region Control events
        private void butExit_Click(object sender, System.EventArgs e)
        {
            Application.Exit();
        }

        private void butRunQuery_Click(object sender, EventArgs e)
        {
            strServer = textServer.Text;
            strUser = textUser.Text;
            strPassword = textPassword.Text;
            strQuery = textMDXQuery.Text;
            strCatalogName = listBoxCubes.SelectedItem.ToString();
            writeConsole();
        }

        private void butViewQuery_Click(object sender, EventArgs e)
        {
            try
            {
                strServer = textServer.Text;
                strUser = textUser.Text;
                strPassword = textPassword.Text;
                strQuery = textMDXQuery.Text;
                strCatalogName = listBoxCubes.SelectedItem.ToString();
                ArrayList queryResult = util.MDXquery(textMDXQuery.Text, connectionString, this.textUser.Text, this.textPassword.Text, strCatalogName);
                CultureInfo _currentCulture = CultureInfo.CurrentCulture;
                _myQuery.arrayResult = queryResult;
                _myQuery.ShowDialog();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Source + " " + ex.Message);
            }
        }

        private void butLogin_Click(object sender, System.EventArgs e)
        {
            connectionString = "Provider=MSOLAP.5;Data Source=" + this.textServer.Text + ";Location=" + this.textServer.Text + ";MDX Compatibility=2";
            saveCatalogName = null;
            listBoxCubes.Items.Clear();
            this.Cursor = Cursors.WaitCursor;

            util.connectionString = connectionString;
            strUser = this.textUser.Text;
            strPassword = this.textPassword.Text;
            try
            {
                butRunQuery.Enabled = false;
                butViewQuery.Enabled = false;
                textMDXQuery.Enabled = false;
                isCubeBuilt = util.LoadCatalogs(out errorMessage);
                if (!String.IsNullOrEmpty(errorMessage))
                    MessageBox.Show(errorMessage);
                if (isCubeBuilt)
                {
                    if (util.listCatalogs.Count > 0)
                    {
                        butRunQuery.Enabled = true;
                        textMDXQuery.Enabled = true;
                        butViewQuery.Enabled = true;
                        DisplayCatalogs(util.listCatalogs);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Source + " " + ex.Message);
            }
            this.Cursor = Cursors.Default;
        }
        #endregion

        #region Write
        private void writeConsole()
        {
            try
            {
                ArrayList queryResult = util.MDXquery(strQuery, connectionString, strUser, strPassword, strCatalogName);
                CultureInfo _currentCulture = CultureInfo.CurrentCulture;
                AllocConsole();
                Console.WriteLine("beginDSInfo");
                Console.WriteLine("csv_first_row_has_column_names;true;true");
                Console.WriteLine("csv_separator;\t;true");
                Console.WriteLine("csv_number_decimal;" + _currentCulture.NumberFormat.NumberDecimalSeparator + ";true");
                Console.WriteLine("csv_date_format;" + _currentCulture.DateTimeFormat.ShortDatePattern + ";true");
                Console.WriteLine("server=" + strServer.Replace("\"", "|~~|") + ";true");
                Console.WriteLine("user=" + strUser.Replace("\"", "|~~|") + ";true");
                Console.WriteLine("password=" + strPassword.Replace("\"", "|~~|") + ";true");
                Console.WriteLine("catalog=" + strCatalogName.Replace("\"", "|~~|") + ";true");
                Console.WriteLine("mdx query=" + strQuery.Replace("\r\n", " ").Replace("\"", "|~~|") + ";true");
                Console.WriteLine("endDSInfo");
                Console.WriteLine("beginData");
                writeData(queryResult);
                Console.WriteLine("endData");
                FreeConsole();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Source + " " + ex.Message);
            }
            this.Close();
        }

        private void writeData(ArrayList arrayResult)
        {
            int nbRows = 0;
            for (int i = 0; i < arrayResult.Count; i++)
            {
                if (nbRows >= size)
                    break;
                string strLine = "";
                string[] str = arrayResult[i].ToString().Split('\t');
                for (int k = 0; k < str.Length; k++)
                {
                    strLine += str[k] + (k == str.Length - 1 ? "" : "\t");
                }
                Console.WriteLine(strLine);
                nbRows++;
            }
        }
        #endregion

        #region Load Catalogs
        private void DisplayCatalogs(ArrayList valueArray)
        {
            string[] strSplit;
            listBoxCubes.Items.Clear();
            bool bSelected = false;
            foreach (Object oCat in valueArray)
            {
                strSplit = oCat.ToString().Split(cArraySplit);
                listBoxCubes.Items.Add(strSplit[0]);
                saveCatalogName = strSplit[0];
                util.strCatalogName = strSplit[0];
                if (saveCatalogName == strCatalogName)
                {
                    bSelected = true;
                    listBoxCubes.SelectedIndex = listBoxCubes.Items.Count - 1;
                }
            }
            if (listBoxCubes.Items.Count > 0)
            {
                if (!bSelected)
                    listBoxCubes.SelectedIndex = 0;
            }
        }
        #endregion
    }
}

