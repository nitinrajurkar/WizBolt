using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.VisualStyles;
using System.IO;
using System.Management;
using System.Web;
using System.Configuration;
using System.Data.SQLite;
using System.Runtime.InteropServices;
using System.Reflection;
using System.Diagnostics;
using System.Drawing.Drawing2D;
using System.Collections;
using System.Text.RegularExpressions;
using WizBolt.Reports;
using System.Security.Cryptography;
using System.Globalization;
using System.Data.OleDb;
using WizBolt.Help;

namespace WizBolt
{
    /// <summary>
    ///   <c>Main</c> 
    ///   Main container that provide frame to Main Form
    /// </summary>
    /// <remarks>
    ///   Commencement of the application
    /// </remarks>
    public partial class WizBoltMainFrame : Form
    {
        static string WindowsDrive = Environment.GetEnvironmentVariable("SystemRoot");  // Finds Windows drive i.e. drive on which Windows O.S. is installed
        public string UserClicked_FileWithPath = string.Empty;

        public WizBoltMainFrame()
        {
            InitializeComponent();
        }

        public WizBoltMainFrame(string FileWithPath)
        {
            InitializeComponent();
            UserClicked_FileWithPath = FileWithPath;
        }

        // Global Variables
        #region Global Variables
        public static Main MainForm;                // = new Main();
        public string EngineerName = string.Empty;
        public DateTime ProjectDate = new DateTime();
        public string SummaryNotes = string.Empty;
        public StringBuilder ProjectNotes = new StringBuilder();
        public decimal CrossLoading = 0.00M;
        public decimal Detension = 0.00M;
        public decimal Friction = 0.00M;
        public string StressBase = string.Empty;
        public string WizBolt_Status = string.Empty;
        public string BasePath = System.IO.Path.GetDirectoryName(Application.ExecutablePath);

       // public static int Application_Id = 0;

       // public string UnitSystem = string.Empty;
        #endregion Global Variables

        // End of Variables
        // Global Functions
        /// <summary>
        ///  Provision for Encryption 
        /// </summary>
        public static string SimpleEncrypt(string IntendedText)
        {
            string EncryptedText = "";

            foreach (char InputChar in IntendedText)
            {
                try
                {
                    uint BinaryOutput = Convert.ToUInt32(InputChar);
                    EncryptedText += (char)(~BinaryOutput);
                }
                catch (Exception)
                {
                    continue;
                }
            }
            return EncryptedText;
        }

        // <summary>  
        /// Encrypts a string           
        /// </summary>           
        /// <param name="PlainText">Text to be encrypted</param>  
        /// <param name="Password">Password to encrypt with</param> 
        /// <param name="Salt">Salt to encrypt with</param> 
        /// <param name="HashAlgorithm">Can be either SHA1 or MD5</param>
        /// <param name="PasswordIterations">Number of iterations to do</param>  
        /// <param name="InitialVector">Needs to be 16 ASCII characters long</param> 
        /// <param name="KeySize">Can be 128, 192, or 256</param>  
        /// <returns>An encrypted string</returns> 
        public static string Encrypt(string PlainText, string Password, string Salt = "Zaniyah", string HashAlgorithm = "SHA1", int PasswordIterations = 5, string InitialVector = "OFRna73m*aze01xY", int KeySize = 256)
        {
            if (string.IsNullOrEmpty(PlainText)) { return ""; }
            byte[] InitialVectorBytes = Encoding.ASCII.GetBytes(InitialVector);
            byte[] SaltValueBytes = Encoding.ASCII.GetBytes(Salt);
            byte[] PlainTextBytes = Encoding.UTF8.GetBytes(PlainText);
            PasswordDeriveBytes DerivedPassword = new PasswordDeriveBytes(Password, SaltValueBytes, HashAlgorithm, PasswordIterations);
            byte[] KeyBytes = DerivedPassword.GetBytes(KeySize / 8);
            RijndaelManaged SymmetricKey = new RijndaelManaged();
            SymmetricKey.Mode = CipherMode.CBC;
            byte[] CipherTextBytes = null;
            using (ICryptoTransform Encryptor = SymmetricKey.CreateEncryptor(KeyBytes, InitialVectorBytes))
            {
                using (MemoryStream MemStream = new MemoryStream())
                {
                    using (CryptoStream CryptoStream = new CryptoStream(MemStream, Encryptor, CryptoStreamMode.Write))
                    {
                        CryptoStream.Write(PlainTextBytes, 0, PlainTextBytes.Length);
                        CryptoStream.FlushFinalBlock();
                        CipherTextBytes = MemStream.ToArray();
                        MemStream.Close();
                        CryptoStream.Close();
                    }
                }
            }
            SymmetricKey.Clear();
            return Convert.ToBase64String(CipherTextBytes);
        }
        /// <summary>  
        /// Decrypts a string  
        /// </summary>  
        /// <param name="CipherText">Text to be decrypted</param>  
        /// <param name="Password">Password to decrypt with</param>  
        /// <param name="Salt">Salt to decrypt with</param>  
        /// <param name="HashAlgorithm">Can be either SHA1 or MD5</param>  
        /// <param name="PasswordIterations">Number of iterations to do</param>  
        /// <param name="InitialVector">Needs to be 16 ASCII characters long</param>  
        /// <param name="KeySize">Can be 128, 192, or 256</param>  
        /// <returns>A decrypted string</returns>  
        public static string Decrypt(string CipherText, string Password, string Salt = "Zaniyah", string HashAlgorithm = "SHA1", int PasswordIterations = 5, string InitialVector = "OFRna73m*aze01xY", int KeySize = 256)
        {
            if (string.IsNullOrEmpty(CipherText)) return "";
            byte[] InitialVectorBytes = Encoding.ASCII.GetBytes(InitialVector);
            byte[] SaltValueBytes = Encoding.ASCII.GetBytes(Salt);
            byte[] CipherTextBytes = Convert.FromBase64String(CipherText);
            PasswordDeriveBytes DerivedPassword = new PasswordDeriveBytes(Password, SaltValueBytes, HashAlgorithm, PasswordIterations);
            byte[] KeyBytes = DerivedPassword.GetBytes(KeySize / 8);
            RijndaelManaged SymmetricKey = new RijndaelManaged();
            SymmetricKey.Mode = CipherMode.CBC;
            byte[] PlainTextBytes = new byte[CipherTextBytes.Length];
            int ByteCount = 0;
            using (ICryptoTransform Decryptor = SymmetricKey.CreateDecryptor(KeyBytes, InitialVectorBytes))
            {
                using (MemoryStream MemStream = new MemoryStream(CipherTextBytes))
                {
                    using (CryptoStream CryptoStream = new CryptoStream(MemStream, Decryptor, CryptoStreamMode.Read))
                    {
                        ByteCount = CryptoStream.Read(PlainTextBytes, 0, PlainTextBytes.Length);
                        MemStream.Close();
                        CryptoStream.Close();
                    }
                }
            }
            SymmetricKey.Clear();
            return Encoding.UTF8.GetString(PlainTextBytes, 0, ByteCount);
        }

        /// <summary>
        /// To gather Hard Disk related information
        /// </summary>
        /// <param name="EntityHDD"></param>
        /// <param name="HDD_Properties"></param>
        /// <returns>Hard Disk Mapped Information</returns>
        public string HDDInfo(string EntityHDD, string HDD_Properties)
        //Return a hardware identifier
        {
            string HDDMap = "";
            System.Management.ManagementClass mc = new System.Management.ManagementClass(EntityHDD);
            System.Management.ManagementObjectCollection moc = mc.GetInstances();
            foreach (System.Management.ManagementObject mo in moc)
            {
                //Only get the first one
                if (HDDMap == "")
                {
                    try
                    {
                        HDDMap = mo[HDD_Properties].ToString();
                        break;
                    }
                    catch
                    {
                    }
                }
            }
            return HDDMap;
        }

        // End of Functions
        /// <summary>
        /// Main Form is loaded based on user selected project file or not
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void WizBoltMainFrame_Load(object sender, EventArgs e)
        {
            if (UserClicked_FileWithPath.Trim() != string.Empty)
            {
                MainForm = new Main(UserClicked_FileWithPath);
                MainForm.MdiParent = this;
                MainForm.Show();
            }
            else
            {
                MainForm = new Main();
                MainForm.MdiParent = this;
                MainForm.Show();
            }
        }

        /// <summary>
        /// Exits the application after performing proper closing tasks.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Exit_ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Control[] TitleControls = MainForm.Controls.Find("ProjectId_Label", true);
            if (Convert.ToInt16(TitleControls[0].Text.Trim()) > 0)
            {
                DialogResult IsPersist = MessageBox.Show("Do you want to save the project?", "User Confirmation!", MessageBoxButtons.YesNoCancel);
                switch (IsPersist)
                {
                    case DialogResult.Yes:
                        MainForm.SaveClose_Project();
                        this.Close();
                        Application.Exit();
                        break;
                    case DialogResult.No:
                        MainForm.CloseProject();
                        this.Close();
                        Application.Exit();
                        break;
                    case DialogResult.Cancel:
                        break;
                }
            }
            else
            {
                this.Close();
                Application.Exit();
            }
        }

        /// <summary>
        /// Exits the application after performing proper closing tasks.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Exit_ToolStripButton_Click(object sender, EventArgs e)
        {
            Control[] TitleControls = MainForm.Controls.Find("ProjectId_Label", true);
            if (Convert.ToInt16(TitleControls[0].Text.Trim()) > 0)
            {
                DialogResult IsPersist = MessageBox.Show("Do you want to save the project?", "User Confirmation!", MessageBoxButtons.YesNoCancel);
                switch (IsPersist)
                {
                    case DialogResult.Yes:
                        MainForm.SaveClose_Project();
                        this.Close();
                        Application.Exit();
                        break;
                    case DialogResult.No:
                        MainForm.CloseProject();
                        this.Close();
                        Application.Exit();
                        break;
                    case DialogResult.Cancel:
                        break;
                }
            }
            else
            {
                this.Close();
                Application.Exit();
            }
        }

        private void ProjectNew_ToolStripButton_Click(object sender, EventArgs e)
        {
            Control[] ProjectControls = MainForm.Controls.Find("ProjectId_Label", true);
            if (Convert.ToInt32(ProjectControls[0].Text.Trim()) > 0)
            {
                DialogResult IsPersist = MessageBox.Show("Do you want to save the project?", "User Confirmation!", MessageBoxButtons.YesNoCancel);
                switch (IsPersist)
                {
                    case DialogResult.Yes:
                        MainForm.SaveClose_Project();
                        Project ProjectForm = new Project();
                        DialogResult ReturnStatus = ProjectForm.ShowDialog();
                        if (ReturnStatus == DialogResult.OK)
                        {
                            PrintAppList_ToolStripButton.Enabled = true;
                            PrintSummary_ToolStripButton.Enabled = true;
                            PrintAppList_ToolStripMenuItem.Enabled = true;
                            PrintSummary_ToolStripMenuItem.Enabled = true;
                        }
                        break;
                    case DialogResult.No:
                        MainForm.CloseProject();
                        Project ProjectDisplay = new Project();
                        DialogResult ProjectStatus = ProjectDisplay.ShowDialog();
                        if (ProjectStatus == DialogResult.OK)
                        {
                            PrintAppList_ToolStripButton.Enabled = true;
                            PrintSummary_ToolStripButton.Enabled = true;
                            PrintAppList_ToolStripMenuItem.Enabled = true;
                            PrintSummary_ToolStripMenuItem.Enabled = true;
                        }
                        break;
                    case DialogResult.Cancel:
                        break;
                }
            }
            else
            {
                Project ProjectForm = new Project();
                DialogResult ReturnStatus = ProjectForm.ShowDialog();
                if (ReturnStatus == DialogResult.OK)
                {
                    PrintAppList_ToolStripButton.Enabled = true;
                    PrintSummary_ToolStripButton.Enabled = true;
                    PrintAppList_ToolStripMenuItem.Enabled = true;
                    PrintSummary_ToolStripMenuItem.Enabled = true;
                }
            }
        }

        private void ProjectNew_ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Control[] ProjectControls = MainForm.Controls.Find("ProjectId_Label", true);
            if (Convert.ToInt32(ProjectControls[0].Text.Trim()) > 0)
            {
                DialogResult IsPersist = MessageBox.Show("Do you want to save the project?", "User Confirmation!", MessageBoxButtons.YesNoCancel);
                switch (IsPersist)
                {
                    case DialogResult.Yes:
                        MainForm.SaveClose_Project();
                        Project ProjectForm = new Project();
                        DialogResult ReturnStatus = ProjectForm.ShowDialog();
                        if (ReturnStatus == DialogResult.OK)
                        {
                            PrintAppList_ToolStripButton.Enabled = true;
                            PrintSummary_ToolStripButton.Enabled = true;
                            PrintAppList_ToolStripMenuItem.Enabled = true;
                            PrintSummary_ToolStripMenuItem.Enabled = true;
                        }
                        break;
                    case DialogResult.No:
                        MainForm.CloseProject();
                        Project ProjectDisplay = new Project();
                        DialogResult ProjectStatus = ProjectDisplay.ShowDialog();
                        if (ProjectStatus == DialogResult.OK)
                        {
                            PrintAppList_ToolStripButton.Enabled = true;
                            PrintSummary_ToolStripButton.Enabled = true;
                            PrintAppList_ToolStripMenuItem.Enabled = true;
                            PrintSummary_ToolStripMenuItem.Enabled = true;
                        }
                        break;
                    case DialogResult.Cancel:
                        break;
                }
            }
            else
            {
                Project ProjectForm = new Project();
                DialogResult ReturnStatus = ProjectForm.ShowDialog();
                if (ReturnStatus == DialogResult.OK)
                {
                    PrintAppList_ToolStripButton.Enabled = true;
                    PrintSummary_ToolStripButton.Enabled = true;
                    PrintAppList_ToolStripMenuItem.Enabled = true;
                    PrintSummary_ToolStripMenuItem.Enabled = true;
                }
            }
        }

        private void SaveNewApplication_ToolStripButton_Click(object sender, EventArgs e)
        {
            MainForm.SaveApplication();
        }

        

        private void SaveAsNewApplication_ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MainForm.SaveApplication();
        }

        private void ProjectViewPage_ToolStripButton_Click(object sender, EventArgs e)
        {
            ReportApp ApplicationReport = new ReportApp();
            ApplicationReport.Show();
        }

        private void PrintGraph_ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ReportApp ApplicationReport = new ReportApp();
            ApplicationReport.Show();
        }

        private void ProjectSave_ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MainForm.SaveProject();
        }

        private void ProjectSave_ToolStripButton_Click(object sender, EventArgs e)
        {
            MainForm.SaveProject();
        }

        private void ProjectOpen_ToolStripButton_Click(object sender, EventArgs e)
        {
            string OpenStatus = MainForm.OpenProject();
            if (OpenStatus == "Successful")
            {
                PrintAppList_ToolStripButton.Enabled = true;
                PrintSummary_ToolStripButton.Enabled = true;
                PrintAppList_ToolStripMenuItem.Enabled = true;
                PrintSummary_ToolStripMenuItem.Enabled = true;
            }
        }

        private void ProjectOpen_ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string OpenStatus = MainForm.OpenProject();
            if (OpenStatus == "Successful")
            {
                PrintAppList_ToolStripButton.Enabled = true;
                PrintSummary_ToolStripButton.Enabled = true;
                PrintAppList_ToolStripMenuItem.Enabled = true;
                PrintSummary_ToolStripMenuItem.Enabled = true;
            }
        }

        private void ModifyProject_ToolStripButton_Click(object sender, EventArgs e)
        {
            string Project_Id = string.Empty;
            Control[] Project_Control = MainForm.Controls.Find("ProjectId_Label", true);
            if (Project_Control.Count() > 0)
            {
                Project_Id = Project_Control[0].Text;

                SQLiteConnection Project_Connection = new SQLiteConnection("Data Source=" + BasePath + "\\WizBolt.db;New=False;");
                SQLiteCommand Project_Command = new SQLiteCommand(Project_Connection);
                Project_Connection.Open();
               
                try
                {
                    Project_Command.CommandText = " SELECT  ProjectReportId,  CustomerName, ProjectName, ProjectReference,  "
                                                    + "  (StartDateDay || '/' || StartDateMonth || '/' || StartDateYear) AS StartDate,    "
                                                    + "  (EndDateDay || '/' || EndDateMonth || '/' || EndDateYear) AS EndDate,  EndDateDay,  "
                                                    + "  EngineerName,  Notes, SummaryNotes, TensionerToolSeriesId, Series,               "
                                                    + "  CrossLoading_PC, Detensioning_PC, Coefficient_of_Friction, StressValue_Base      "
                                                    + "  	FROM ProjectReport                                                            "
                                                    + "  			WHERE  ProjectReportId > 0 AND ProjectReportId = " + Project_Id;

                    DataSet UpdateProject_DataSet = new DataSet();
                    SQLiteDataAdapter UpdateProject_Adapter = new SQLiteDataAdapter(Project_Command);

                    int Project_Result = UpdateProject_Adapter.Fill(UpdateProject_DataSet);
                    if (Project_Result > 0)
                    {
                        ProjectView CurrentProject = new ProjectView();
                        CurrentProject.ProjectDisplayId = Project_Id;
                        CurrentProject.ShowDialog();
                    }
                    else
                    {
                        MessageBox.Show("Project not found. Either there may not be any current project or project is not saved. If there is no project, either upload existing project or create new", "Error!");
                    }
                }
                catch (Exception ProjectOpen_Exception)
                {
                    MessageBox.Show( ProjectOpen_Exception.Message + " while verifying project existence.", "Fatal Error!!!");
                }
            }
        }

        private void ModifyProject_ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string Project_Id = string.Empty;
            Control[] Project_Control = MainForm.Controls.Find("ProjectId_Label", true);
            if (Project_Control.Count() > 0)
            {
                Project_Id = Project_Control[0].Text;

                SQLiteConnection Project_Connection = new SQLiteConnection("Data Source=" + BasePath + "\\WizBolt.db;New=False;");
                SQLiteCommand Project_Command = new SQLiteCommand(Project_Connection);
                Project_Connection.Open();

                try
                {
                    Project_Command.CommandText = " SELECT  ProjectReportId,  CustomerName, ProjectName, ProjectReference,  "
                                                    + "  (StartDateDay || '/' || StartDateMonth || '/' || StartDateYear) AS StartDate,    "
                                                    + "  (EndDateDay || '/' || EndDateMonth || '/' || EndDateYear) AS EndDate,  EndDateDay,  "
                                                    + "  EngineerName,  Notes, SummaryNotes, TensionerToolSeriesId, Series,               "
                                                    + "  CrossLoading_PC, Detensioning_PC, Coefficient_of_Friction, StressValue_Base      "
                                                    + "  	FROM ProjectReport                                                            "
                                                    + "  			WHERE  ProjectReportId > 0 AND ProjectReportId = " + Project_Id;

                    DataSet UpdateProject_DataSet = new DataSet();
                    SQLiteDataAdapter UpdateProject_Adapter = new SQLiteDataAdapter(Project_Command);

                    int Project_Result = UpdateProject_Adapter.Fill(UpdateProject_DataSet);
                    if (Project_Result > 0)
                    {
                        ProjectView CurrentProject = new ProjectView();
                        CurrentProject.ProjectDisplayId = Project_Id;
                        CurrentProject.ShowDialog();
                    }
                    else
                    {
                        MessageBox.Show("Project not found. Either there may not be any current project or project is not saved. If there is no project, either upload existing project or create new", "Error!");
                    }
                }
                catch (Exception ProjectOpen_Exception)
                {
                    MessageBox.Show(ProjectOpen_Exception.Message + " while verifying project existence.", "Fatal Error!!!");
                }
            }
        }

        private void SaveApplication_ToolStripButton_Click(object sender, EventArgs e)
        {
            MainForm.UpdateApplication();
        }

        private void saveApplicationToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MainForm.UpdateApplication();
        }

        private void ProjectDelete_ToolStripButton_Click(object sender, EventArgs e)
        {
            //DialogResult UserResponse = MessageBox.Show("This project will be parmanently deleted and you will never be able to get its contents again. This is irreversible process. Do you really want to delete?", "Warning!", MessageBoxButtons.YesNo);
            //if (UserResponse == DialogResult.Yes)
            //{
                MainForm.RemoveProject();
            //}
        }

        private void ProjectDelete_ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            
                MainForm.RemoveProject();
            
        }

        private void ApplicationDelete_ToolStripButton_Click(object sender, EventArgs e)
        {
            DialogResult UserAct = MessageBox.Show("This application will be parmanently deleted and you will never be able to get its contents again. This is irreversible process. Do you really want to delete?", "Warning!", MessageBoxButtons.YesNo);
            if (UserAct == DialogResult.Yes)
            {
                MainForm.RemoveApplication();
            }
        }

        private void ApplicationDelete_ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DialogResult UserAct = MessageBox.Show("This application will be parmanently deleted and you will never be able to get its contents again. This is irreversible process. Do you really want to delete?", "Warning!", MessageBoxButtons.YesNo);
            if (UserAct == DialogResult.Yes)
            {
                MainForm.RemoveApplication();
            }
        }

        private void ProjectClose_ToolStripButton_Click(object sender, EventArgs e)
        {
            Control[] TitleControls = MainForm.Controls.Find("ProjectId_Label", true);
            if (Convert.ToInt16(TitleControls[0].Text.Trim()) > 0)
            {
                DialogResult UserInput = MessageBox.Show("Do you want to save and then close (Yes)  or close without saving (No) or cancel?", "Warning!", MessageBoxButtons.YesNoCancel);
                if (UserInput == DialogResult.Yes)
                {
                    MainForm.SaveClose_Project();
                }
                else if (UserInput == DialogResult.No)
                {
                    MainForm.CloseProject();
                }
            }
            else
            {
                MessageBox.Show("No project is open.", "User Information");
            }
        }

        private void ProjectClose_ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Control[] TitleControls = MainForm.Controls.Find("ProjectId_Label", true);
            if (Convert.ToInt16(TitleControls[0].Text.Trim()) > 0)
            {
                DialogResult UserInput = MessageBox.Show("Do you want to save and then close (Yes) or close without saving (No) or cancel?", "Warning!", MessageBoxButtons.YesNoCancel);
                if (UserInput == DialogResult.Yes)
                {
                    MainForm.SaveClose_Project();
                }
                else if (UserInput == DialogResult.No)
                {
                    MainForm.CloseProject();
                }
            }
            else
            {
                MessageBox.Show("No project is open.", "User Information");
            }
        }

        private void SaveAs_ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //ReportApp ApplicationReport = new ReportApp();
            //ApplicationReport.Show();
            MainForm.SaveAs();          // Saves project with different name
        }

        private void PrintAppList_ToolStripButton_Click(object sender, EventArgs e)
        {
            ApplicationList AppList_Report = new ApplicationList();
            AppList_Report.Show();
        }

        private void PrintAppList_ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ApplicationList AppList_Report = new ApplicationList();
            AppList_Report.Show();
        }

        private void PrintSummary_ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SummaryReport Summary_Report = new SummaryReport();
            Summary_Report.Show();
        }

        private void PrintSummary_ToolStripButton_Click(object sender, EventArgs e)
        {
            SummaryReport Summary_Report = new SummaryReport();
            Summary_Report.Show();
        }

        private void SpecialTools_ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MainForm.SpecialTools();
        }

        protected DataSet ManageData()
        {
            DataSet Wiz_DataSet = new DataSet();
            string DataFile = string.Empty;
            string FileExtension = string.Empty;
            string ExcelConnect = string.Empty;
            string HDR = "Yes";             // Excel file has column names in the first row.

            OpenFileDialog OpenExcelData = new OpenFileDialog();
            OpenExcelData.Filter = "Excel files (*.xlsx)|*.xlsx|Excel Macro (*.xlsm)|*.xlsm|Excel Old (*.xls)|*.xls";
            OpenExcelData.RestoreDirectory = true;
            OpenExcelData.Multiselect = false;
            DialogResult OpenDataResult = OpenExcelData.ShowDialog();
            if (OpenDataResult == DialogResult.OK)
            {
                DataFile = OpenExcelData.FileName;
                if (File.Exists(DataFile))
                {
                    FileExtension = DataFile.Substring(DataFile.Length - 3, 3);
                    if (FileExtension == "xls")                // Old Excel file format i.e. upto Excel 2003.
                    {
                        ExcelConnect = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + DataFile + ";Extended Properties='Excel 8.0;HDR=" + HDR + ";IMEX=1'";
                    }
                    else                                      // New Excel file format i.e. from Excel 2007 onwards.
                    {
                        ExcelConnect = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + DataFile + ";Extended Properties=\"Excel 12.0 Xml;HDR=" + HDR + ";IMEX=1;\"";
                    }

                    using (OleDbConnection conn = new OleDbConnection(ExcelConnect))
                    {
                        conn.Open();

                        DataTable Excel_DataTable = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });
                        int SheetCount = 0;
                        foreach (DataRow row in Excel_DataTable.Rows)
                        {
                            ++SheetCount;
                            string sheet = row["TABLE_NAME"].ToString();

                            OleDbCommand cmd = new OleDbCommand("SELECT * FROM [" + sheet + "]", conn);
                            cmd.CommandType = CommandType.Text;

                            DataTable Worksheet_DataTable = new DataTable(sheet);
                            Wiz_DataSet.Tables.Add(Worksheet_DataTable);

                            new OleDbDataAdapter(cmd).Fill(Worksheet_DataTable);
                            if (SheetCount == 1)
                            {
                                Worksheet_DataTable.Rows[0].Delete();
                                Worksheet_DataTable.AcceptChanges();                    // This is must after delete as the index do not change and delete information remains giving error "Deleted row information cannot be accessed through the row".
                            }
                        }
                    }
                }
            }

            return Wiz_DataSet;
        }

        private void BoltData_ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string Insert_Status = "Successful";
            DataSet Bolt_DataSet = ManageData();
            if ((Bolt_DataSet != null) && (Bolt_DataSet.Tables.Count > 0))
            {
                DataTable Bolt_DataTable = Bolt_DataSet.Tables[0];

                SQLiteConnection Data_Connection = new SQLiteConnection("Data Source=" + BasePath + "\\WizBolt.db;New=False;");
                SQLiteCommand Data_SQLiteCommand = new SQLiteCommand(Data_Connection);
                Data_Connection.Open();

                foreach (DataRow BoltRow in Bolt_DataTable.Rows)
                {
                    Data_SQLiteCommand.CommandText = "  INSERT INTO Bolt (StandardId, StandardAuxiliaryId, MaterialId, BoltId, BoltAbbreviation, BoltName, BoltAlternateName, Bolt_ImperialName, Bolt_SIName, BoltSize,	Default_UnitSystemId, ThreadType, Threads_Per_Inch, Pitch, BoltLength, "    // 15
                                                      + " NominalDiameter,  MinimumDiameter, MaximumDiameter, TensileStressArea, TensileStress, Yield, Elongation, Reduction_in_Area, Hardness) "   // 24
                                                      + "  VALUES (" + BoltRow.ItemArray[0] + ", "
                                                      + BoltRow.ItemArray[1] + ", "
                                                      + BoltRow.ItemArray[2] + ", "
                                                      + BoltRow.ItemArray[3] + ", "
                                                      + "'" + BoltRow.ItemArray[4] + "', "
                                                      + "'" + BoltRow.ItemArray[5] + "', "
                                                      + "'" + BoltRow.ItemArray[6] + "', "
                                                      + "'" + BoltRow.ItemArray[7] + "', "
                                                      + "'" + BoltRow.ItemArray[8] + "', "
                                                      + BoltRow.ItemArray[9] + ", "
                                                      + BoltRow.ItemArray[10] + ", "
                                                      + "'" + BoltRow.ItemArray[11] + "', "
                                                      + BoltRow.ItemArray[12] + ", "
                                                      + BoltRow.ItemArray[13] + ", "
                                                      + BoltRow.ItemArray[14] + ", "
                                                      + BoltRow.ItemArray[15] + ", "
                                                      + BoltRow.ItemArray[16] + ", "
                                                      + BoltRow.ItemArray[17] + ", "
                                                      + BoltRow.ItemArray[18] + ", "
                                                      + BoltRow.ItemArray[19] + ", "
                                                      + BoltRow.ItemArray[20] + ", "
                                                      + BoltRow.ItemArray[21] + ", "
                                                      + BoltRow.ItemArray[22] + ", "
                                                      + BoltRow.ItemArray[23] + ") ";
                    try
                    {
                        int BoltResult = Data_SQLiteCommand.ExecuteNonQuery();
                    }
                    catch (Exception Bolt_Exception)
                    {
                        MessageBox.Show(Bolt_Exception + " while adding new data from Excel file in Bolt table.");
                        Insert_Status = "Failed";
                    }
                }
                if (Insert_Status == "Successful")
                {
                    MessageBox.Show("Bolt data uploaded successfully.", "Information for User!");
                }
            }
        }

        private void BoltMaterialData_ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string Insert_Status = "Successful";
            DataSet BoltMaterial_DataSet = ManageData();
            if ((BoltMaterial_DataSet != null) && (BoltMaterial_DataSet.Tables.Count > 0))
            {
                DataTable BoltMaterial_DataTable = BoltMaterial_DataSet.Tables[0];

                SQLiteConnection Data_Connection = new SQLiteConnection("Data Source=" + BasePath + "\\WizBolt.db;New=False;");
                SQLiteCommand Data_SQLiteCommand = new SQLiteCommand(Data_Connection);
                Data_Connection.Open();

                foreach (DataRow BoltRow in BoltMaterial_DataTable.Rows)
                {
                    Data_SQLiteCommand.CommandText = "  INSERT INTO BoltMaterial (MaterialId, BoltMaterialId, Min_BoltDiameter, Max_BoltDiameter, TensileStress, Elongation, Reduction_in_Area, Yield, Hardness) "
                                                      + "  VALUES (" + BoltRow.ItemArray[0] + ", "
                                                      + BoltRow.ItemArray[1] + ", "
                                                      + BoltRow.ItemArray[2] + ", "
                                                      + BoltRow.ItemArray[3] + ", "
                                                      + BoltRow.ItemArray[4] + ", "
                                                      + BoltRow.ItemArray[5] + ", "
                                                      + BoltRow.ItemArray[6] + ", "
                                                      + BoltRow.ItemArray[7] + ", "
                                                      + BoltRow.ItemArray[8] + ") ";
                    try
                    {
                        int BoltResult = Data_SQLiteCommand.ExecuteNonQuery();
                    }
                    catch (Exception Bolt_Exception)
                    {
                        MessageBox.Show(Bolt_Exception + " while adding new data from Excel file in Bolt Materials table.");
                        Insert_Status = "Failed";
                    }
                }
                if (Insert_Status == "Successful")
                {
                    MessageBox.Show("Bolt material data uploaded successfully.", "Information for User!");
                }
                Data_Connection.Close();
                Data_SQLiteCommand.Dispose();
                Data_Connection.Dispose();
            }
        }

        private void FlangeTypesData_ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string Insert_Status = "Successful";
            DataSet FlangeTypes_DataSet = ManageData();
            if ((FlangeTypes_DataSet != null) && (FlangeTypes_DataSet.Tables.Count > 0))
            {
                DataTable FlangeTypes_DataTable = FlangeTypes_DataSet.Tables[0];

                SQLiteConnection Data_Connection = new SQLiteConnection("Data Source=" + BasePath + "\\WizBolt.db;New=False;");
                SQLiteCommand Data_SQLiteCommand = new SQLiteCommand(Data_Connection);
                Data_Connection.Open();

                foreach (DataRow FlangeRow in FlangeTypes_DataTable.Rows)
                {
                    Data_SQLiteCommand.CommandText = "  INSERT INTO FlangeTypes (FlangeTypeId, FlangeType, FlangeFace, FlangeAbbreviation) "
                                                      + "  VALUES (" + FlangeRow.ItemArray[0] + ", "
                                                      + "'" + FlangeRow.ItemArray[1] + "', "
                                                      + "'" + FlangeRow.ItemArray[2] + "', "
                                                      + "'" + FlangeRow.ItemArray[3] + "') ";
                    try
                    {
                        int FlangeResult = Data_SQLiteCommand.ExecuteNonQuery();
                    }
                    catch (Exception Flange_Exception)
                    {
                        MessageBox.Show(Flange_Exception + " while adding new data from Excel file in Flange Types table.");
                        Insert_Status = "Failed";
                    }
                }
                if (Insert_Status == "Successful")
                {
                    MessageBox.Show("Flange types data uploaded successfully.", "Information for User!");
                }
                Data_Connection.Close();
                Data_SQLiteCommand.Dispose();
                Data_Connection.Dispose();
            }
        }

        private void FlangeData_ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string Insert_Status = "Successful";
            DataSet Flanges_DataSet = ManageData();
            if ((Flanges_DataSet != null) && (Flanges_DataSet.Tables.Count > 0))
            {
                DataTable Flanges_DataTable = Flanges_DataSet.Tables[0];

                SQLiteConnection Data_Connection = new SQLiteConnection("Data Source=" + BasePath + "\\WizBolt.db;New=False;");
                SQLiteCommand Data_SQLiteCommand = new SQLiteCommand(Data_Connection);
                Data_Connection.Open();

                foreach (DataRow FlangesRow in Flanges_DataTable.Rows)
                {
                    Data_SQLiteCommand.CommandText = "  INSERT INTO Flanges (StandardId, StandardAuxiliaryId, FlangeRatingId, FlangeId, FlangeAbbreviation, "
                                                    + " FlangeName, FlangeAlternateName, FlangeSize, FlangeTypeId, NominalSize, FlangeThickness, SealRingGasketThickness, Number_of_Bolt_Holes, "
                                                    + " Diameter_of_Bolt_Holes,	Bolt_Diameter, ResidualBoltStress) "        // 18
                                                      + "  VALUES (" + FlangesRow.ItemArray[0] + ", "
                                                      + FlangesRow.ItemArray[1] + ", "
                                                      + FlangesRow.ItemArray[2] + ", "
                                                      + FlangesRow.ItemArray[3] + ", "
                                                      + "'" + FlangesRow.ItemArray[4] + "', "
                                                      + "'" + FlangesRow.ItemArray[5] + "', "
                                                      + "'" + FlangesRow.ItemArray[6] + "', "
                                                      + "'" + FlangesRow.ItemArray[7] + "', "
                                                      + FlangesRow.ItemArray[8] + ", "
                                                      + FlangesRow.ItemArray[9] + ", "
                                                      + FlangesRow.ItemArray[10] + ", "
                                                      + FlangesRow.ItemArray[11] + ", "
                                                      + FlangesRow.ItemArray[12] + ", "
                                                      + FlangesRow.ItemArray[13] + ", "
                                                      + FlangesRow.ItemArray[14] + ", "
                                                      + FlangesRow.ItemArray[15]
                                                      + ") ";
                    try
                    {
                        int FlangesResult = Data_SQLiteCommand.ExecuteNonQuery();
                    }
                    catch (Exception Flanges_Exception)
                    {
                        MessageBox.Show(Flanges_Exception + " while adding new data from Excel file in Flanges table.");
                        Insert_Status = "Failed";
                    }
                }
                if (Insert_Status == "Successful")
                {
                    MessageBox.Show("Flanges data uploaded successfully.", "Information for User!");
                }
                Data_Connection.Close();
                Data_SQLiteCommand.Dispose();
                Data_Connection.Dispose();
            }
        }

        private void FlangeRatingData_ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string Insert_Status = "Successful";
            DataSet FlangeRatings_DataSet = ManageData();
            if ((FlangeRatings_DataSet != null) && (FlangeRatings_DataSet.Tables.Count > 0))
            {
                DataTable FlangeRatings_DataTable = FlangeRatings_DataSet.Tables[0];

                SQLiteConnection Data_Connection = new SQLiteConnection("Data Source=" + BasePath + "\\WizBolt.db;New=False;");
                SQLiteCommand Data_SQLiteCommand = new SQLiteCommand(Data_Connection);
                Data_Connection.Open();

                foreach (DataRow RatingRow in FlangeRatings_DataTable.Rows)
                {
                    Data_SQLiteCommand.CommandText = "  INSERT INTO FlangeRatings (StandardId, StandardAuxiliaryId, FlangeRatingId, FlangeRatingClass, NPS, NPSName, FlangeRatingName, AbbreviatedRating, NominalPipeSize) "
                                                      + "  VALUES (" + RatingRow.ItemArray[0] + ", "
                                                      + RatingRow.ItemArray[1] + ", "
                                                      + RatingRow.ItemArray[2] + ", "
                                                      + "'" + RatingRow.ItemArray[3] + "', "
                                                      + "'" + RatingRow.ItemArray[4] + "', "
                                                      + "'" + RatingRow.ItemArray[5] + "', "
                                                      + "'" + RatingRow.ItemArray[6] + "', "
                                                      + "'" + RatingRow.ItemArray[7] + "', "
                                                      + RatingRow.ItemArray[8] + ") ";
                    try
                    {
                        int RatingResult = Data_SQLiteCommand.ExecuteNonQuery();
                    }
                    catch (Exception Rating_Exception)
                    {
                        MessageBox.Show(Rating_Exception + " while adding new data from Excel file in Flange Ratings table.");
                        Insert_Status = "Failed";
                    }
                }
                if (Insert_Status == "Successful")
                {
                    MessageBox.Show("Data uploaded successfully.", "Information for User!");
                }
                Data_Connection.Close();
                Data_SQLiteCommand.Dispose();
                Data_Connection.Dispose();
            }
        }

        private void FlangeBoltData_ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string Insert_Status = "Successful";
            DataSet FlangeBolts_DataSet = ManageData();
            if ((FlangeBolts_DataSet != null) && (FlangeBolts_DataSet.Tables.Count > 0))
            {
                DataTable FlangeBolts_DataTable = FlangeBolts_DataSet.Tables[0];

                SQLiteConnection Data_Connection = new SQLiteConnection("Data Source=" + BasePath + "\\WizBolt.db;New=False;");
                SQLiteCommand Data_SQLiteCommand = new SQLiteCommand(Data_Connection);
                Data_Connection.Open();

                foreach (DataRow BoltsRow in FlangeBolts_DataTable.Rows)
                {
                    Data_SQLiteCommand.CommandText = "  INSERT INTO FlangeBolts (StandardId, StandardAuxiliaryId, FlangeRatingId, FlangeId, FlangeBoltId, "
                                                      + " BoltThread, TPI_Pitch, BoltLength) "
                                                      + "  VALUES (" + BoltsRow.ItemArray[0] + ", "
                                                      + BoltsRow.ItemArray[1] + ", "
                                                      + BoltsRow.ItemArray[2] + ", "
                                                      + BoltsRow.ItemArray[3] + ", "
                                                      + BoltsRow.ItemArray[4] + ", "
                                                      + "'" + BoltsRow.ItemArray[5] + "', "
                                                      + BoltsRow.ItemArray[6] + ", "
                                                      + BoltsRow.ItemArray[7] + ") ";
                    try
                    {
                        int BoltsResult = Data_SQLiteCommand.ExecuteNonQuery();
                    }
                    catch (Exception Bolts_Exception)
                    {
                        MessageBox.Show(Bolts_Exception + " while adding new data from Excel file in Flange bolts table.");
                        Insert_Status = "Failed";
                    }
                }
                if (Insert_Status == "Successful")
                {
                    MessageBox.Show("Flange bolts data uploaded successfully.", "Information for User!");
                }
                Data_Connection.Close();
                Data_SQLiteCommand.Dispose();
                Data_Connection.Dispose();
            }
        }

        private void FlangeToolData_ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string Insert_Status = "Successful";
            DataSet FlangeTools_DataSet = ManageData();
            if ((FlangeTools_DataSet != null) && (FlangeTools_DataSet.Tables.Count > 0))
            {
                DataTable FlangeTools_DataTable = FlangeTools_DataSet.Tables[0];

                SQLiteConnection Data_Connection = new SQLiteConnection("Data Source=" + BasePath + "\\WizBolt.db;New=False;");
                SQLiteCommand Data_SQLiteCommand = new SQLiteCommand(Data_Connection);
                Data_Connection.Open();

                foreach (DataRow ToolsRow in FlangeTools_DataTable.Rows)
                {
                    Data_SQLiteCommand.CommandText = "  INSERT INTO FlangeBoltTools (StandardId, StandardAuxiliaryId, FlangeRatingId, FlangeId, FlangeBoltId, TensionerToolSeriesId, TensionerToolId) "
                                                      + "  VALUES (" + ToolsRow.ItemArray[0] + ", "
                                                      + ToolsRow.ItemArray[1] + ", "
                                                      + ToolsRow.ItemArray[2] + ", "
                                                      + ToolsRow.ItemArray[3] + ", "
                                                      + ToolsRow.ItemArray[4] + ", "
                                                      + ToolsRow.ItemArray[5] + ", "
                                                      + ToolsRow.ItemArray[6] + ") ";
                    try
                    {
                        int ToolsResult = Data_SQLiteCommand.ExecuteNonQuery();
                    }
                    catch (Exception Tools_Exception)
                    {
                        MessageBox.Show(Tools_Exception + " while adding new data from Excel file in Flange bolt tools table.");
                        Insert_Status = "Failed";
                    }
                }
                if (Insert_Status == "Successful")
                {
                    MessageBox.Show("Flange bolt tools data uploaded successfully.", "Information for User!");
                }
                Data_Connection.Close();
                Data_SQLiteCommand.Dispose();
                Data_Connection.Dispose();
            }
        }

        private void MaterialData_ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string Insert_Status = "Successful";
            DataSet Materials_DataSet = ManageData();
            if ((Materials_DataSet != null) && (Materials_DataSet.Tables.Count > 0))
            {
                DataTable Materials_DataTable = Materials_DataSet.Tables[0];

                SQLiteConnection Data_Connection = new SQLiteConnection("Data Source=" + BasePath + "\\WizBolt.db;New=False;");
                SQLiteCommand Data_SQLiteCommand = new SQLiteCommand(Data_Connection);
                Data_Connection.Open();

                foreach (DataRow MaterialRow in Materials_DataTable.Rows)
                {
                    Data_SQLiteCommand.CommandText = "  INSERT INTO Materials (StandardId, StandardAuxiliaryId, MaterialId, MaterialName, MaterialAbbreviation, "
                                                     + "  MaterialAlternateName, MaterialDescription, Alloy_Compound, MaximumStrength, MinimumStrength, TensileStress, Elongation, Reduction_in_Area, Yield) "
                                                      + "  VALUES (" + MaterialRow.ItemArray[0] + ", "
                                                      + MaterialRow.ItemArray[1] + ", "
                                                      + MaterialRow.ItemArray[2] + ", "
                                                      + "'" + MaterialRow.ItemArray[3] + "', "
                                                      + "'" + MaterialRow.ItemArray[4] + "', "
                                                      + "'" + MaterialRow.ItemArray[5] + "', "
                                                      + "'" + MaterialRow.ItemArray[6] + "', "
                                                      + "'" + MaterialRow.ItemArray[7] + "', "
                                                      + MaterialRow.ItemArray[8] + ", "
                                                      + MaterialRow.ItemArray[9] + ", "
                                                      + MaterialRow.ItemArray[10] + ", "
                                                      + MaterialRow.ItemArray[11] + ", "
                                                      + MaterialRow.ItemArray[12] + ", "
                                                      + MaterialRow.ItemArray[13]
                                                      + ") ";
                    try
                    {
                        int MaterialResult = Data_SQLiteCommand.ExecuteNonQuery();
                    }
                    catch (Exception Material_Exception)
                    {
                        MessageBox.Show(Material_Exception + " while adding new data from Excel file in Materials table.");
                        Insert_Status = "Failed";
                    }
                }
                if (Insert_Status == "Successful")
                {
                    MessageBox.Show("Materials data uploaded successfully.", "Information for User!");
                }
                Data_Connection.Close();
                Data_SQLiteCommand.Dispose();
                Data_Connection.Dispose();
            }
        }

        private void Standards_ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string Insert_Status = "Successful";
            DataSet Standards_DataSet = ManageData();
            if ((Standards_DataSet != null) && (Standards_DataSet.Tables.Count > 0))
            {
                DataTable Standards_DataTable = Standards_DataSet.Tables[0];

                SQLiteConnection Data_Connection = new SQLiteConnection("Data Source=" + BasePath + "\\WizBolt.db;New=False;");
                SQLiteCommand Data_SQLiteCommand = new SQLiteCommand(Data_Connection);
                Data_Connection.Open();

                foreach (DataRow StandardRow in Standards_DataTable.Rows)
                {
                    Data_SQLiteCommand.CommandText = "  INSERT INTO Standards (StandardId, StandardAbbreviation, StandardName, StandardAlternateName, StandardDescription, Specifications) "
                                                      + "  VALUES (" + StandardRow.ItemArray[0] + ", "
                                                      + "'" + StandardRow.ItemArray[1] + "', "
                                                      + "'" + StandardRow.ItemArray[2] + "', "
                                                      + "'" + StandardRow.ItemArray[3] + "', "
                                                      + "'" + StandardRow.ItemArray[4] + "', "
                                                      + "'" + StandardRow.ItemArray[5]
                                                      + "') ";
                    try
                    {
                        int StandardResult = Data_SQLiteCommand.ExecuteNonQuery();
                    }
                    catch (Exception Standard_Exception)
                    {
                        MessageBox.Show(Standard_Exception + " while adding new data from Excel file in Standards table.");
                        Insert_Status = "Failed";
                    }
                }
                if (Insert_Status == "Successful")
                {
                    MessageBox.Show("Standards data uploaded successfully.", "Information for User!");
                }
                Data_Connection.Close();
                Data_SQLiteCommand.Dispose();
                Data_Connection.Dispose();
            }
        }

        private void StandardAuxiliaries_ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string Insert_Status = "Successful";
            DataSet StandardAuxiliaries_DataSet = ManageData();
            if ((StandardAuxiliaries_DataSet != null) && (StandardAuxiliaries_DataSet.Tables.Count > 0))
            {
                DataTable StandardAuxiliaries_DataTable = StandardAuxiliaries_DataSet.Tables[0];

                SQLiteConnection Data_Connection = new SQLiteConnection("Data Source=" + BasePath + "\\WizBolt.db;New=False;");
                SQLiteCommand Data_SQLiteCommand = new SQLiteCommand(Data_Connection);
                Data_Connection.Open();

                foreach (DataRow AuxiliaryRow in StandardAuxiliaries_DataTable.Rows)
                {
                    Data_SQLiteCommand.CommandText = "  INSERT INTO StandardAuxiliaries (StandardId, StandardAuxiliaryId, StandardAuxiliary, StandardShortAuxiliary, AuxiliaryAlternateName, "
                                                      + " AuxiliaryDescription, AuxiliaryApplicability, Default_UnitSystemId) "
                                                      + "  VALUES (" + AuxiliaryRow.ItemArray[0] + ", "
                                                      + AuxiliaryRow.ItemArray[1] + ", "
                                                      + "'" + AuxiliaryRow.ItemArray[2] + "', "
                                                      + "'" + AuxiliaryRow.ItemArray[3] + "', "
                                                      + "'" + AuxiliaryRow.ItemArray[4] + "', "
                                                      + "'" + AuxiliaryRow.ItemArray[5] + "', "
                                                      + "'" + AuxiliaryRow.ItemArray[6] + "', "
                                                      + AuxiliaryRow.ItemArray[7]
                                                      + ") ";
                    try
                    {
                        int AuxiliaryResult = Data_SQLiteCommand.ExecuteNonQuery();
                    }
                    catch (Exception Auxiliary_Exception)
                    {
                        MessageBox.Show(Auxiliary_Exception + " while adding new data from Excel file in Auxiliary standards table.");
                        Insert_Status = "Failed";
                    }
                }
                if (Insert_Status == "Successful")
                {
                    MessageBox.Show("Auxiliary standards data uploaded successfully.", "Information for User!");
                }
                Data_Connection.Close();
                Data_SQLiteCommand.Dispose();
                Data_Connection.Dispose();
            }
        }

        private void TensionerToolSeriesData_ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string Insert_Status = "Successful";
            DataSet TensionerToolSeries_DataSet = ManageData();
            if ((TensionerToolSeries_DataSet != null) && (TensionerToolSeries_DataSet.Tables.Count > 0))
            {
                DataTable TensionerToolSeries_DataTable = TensionerToolSeries_DataSet.Tables[0];

                SQLiteConnection Data_Connection = new SQLiteConnection("Data Source=" + BasePath + "\\WizBolt.db;New=False;");
                SQLiteCommand Data_SQLiteCommand = new SQLiteCommand(Data_Connection);
                Data_Connection.Open();

                foreach (DataRow SeriesRow in TensionerToolSeries_DataTable.Rows)
                {
                    Data_SQLiteCommand.CommandText = "  INSERT INTO TensionerTools_Series (TensionerToolSeriesId, Series, Description, AppliedField) "
                                                      + "  VALUES (" + SeriesRow.ItemArray[0] + ", "
                                                      + "'" + SeriesRow.ItemArray[1] + "', "
                                                      + "'" + SeriesRow.ItemArray[2] + "', "
                                                      + "'" + SeriesRow.ItemArray[3] + "') ";
                    try
                    {
                        int SeriesResult = Data_SQLiteCommand.ExecuteNonQuery();
                    }
                    catch (Exception Series_Exception)
                    {
                        MessageBox.Show(Series_Exception + " while adding new data from Excel file in Tensioner tool series table.");
                        Insert_Status = "Failed";
                    }
                }
                if (Insert_Status == "Successful")
                {
                    MessageBox.Show("Tensioner tool series data uploaded successfully.", "Information for User!");
                }
                Data_Connection.Close();
                Data_SQLiteCommand.Dispose();
                Data_Connection.Dispose();
            }
        }

        private void TensionerToolData_ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string Insert_Status = "Successful";
            DataSet TensionerTools_DataSet = ManageData();
            if ((TensionerTools_DataSet != null) && (TensionerTools_DataSet.Tables.Count > 0))
            {
                DataTable TensionerTools_DataTable = TensionerTools_DataSet.Tables[0];

                SQLiteConnection Data_Connection = new SQLiteConnection("Data Source=" + BasePath + "\\WizBolt.db;New=False;");
                SQLiteCommand Data_SQLiteCommand = new SQLiteCommand(Data_Connection);
                Data_Connection.Open();

                foreach (DataRow ToolRow in TensionerTools_DataTable.Rows)
                {
                    Data_SQLiteCommand.CommandText = "  INSERT INTO TensionerTools (TensionerToolSeriesId, TensionerToolId, ModelNumber, MaximumLoad, HydraulicArea, OuterDiameter) "
                                                      + "  VALUES (" + ToolRow.ItemArray[0] + ", "
                                                      + ToolRow.ItemArray[1] + ", "
                                                      + "'" + ToolRow.ItemArray[2] + "', "
                                                      + ToolRow.ItemArray[3] + ", "
                                                      + ToolRow.ItemArray[4] + ", "
                                                      + ToolRow.ItemArray[5] + ") ";
                    try
                    {
                        int ToolResult = Data_SQLiteCommand.ExecuteNonQuery();
                    }
                    catch (Exception Tool_Exception)
                    {
                        MessageBox.Show(Tool_Exception + " while adding new data from Excel file in Tools table.");
                        Insert_Status = "Failed";
                    }
                }
                if (Insert_Status == "Successful")
                {
                    MessageBox.Show("Tools data uploaded successfully.", "Information for User!");
                }
                Data_Connection.Close();
                Data_SQLiteCommand.Dispose();
                Data_Connection.Dispose();
            }
        }

        private void aboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            About_WizBolt About_Window = new About_WizBolt();
            About_Window.MdiParent = this;
            About_Window.Show();
        }

    }
}
