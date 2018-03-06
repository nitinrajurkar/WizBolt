using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
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
using System.Linq;
using System.Collections;
using System.Text.RegularExpressions;
using System.Windows.Forms.DataVisualization.Charting;
using System.Data.OleDb;
using WizBolt.SpecialTools;
using WizBolt.Reports;
using System.Security.Cryptography;

namespace WizBolt
{
    public partial class Main : Form
    {
        // Global Parameters
        #region Global Variables
            //static DAL Dal_Obj = new DAL();         // Data Access Layer (DAL) 
            public bool IsInitial = true;
            public bool IsLoading = false;
            public bool IsSetting = false;
            public bool IsRating = false;
            public bool IsBoltThread = false;
            public decimal LTF_MainCurrent = 0M;
            public DateTime Today = DateTime.Now;
            public string ProjectSavedFile = string.Empty;
            public string ProjectOpenedFile = string.Empty;
            public decimal GlobalMinorDiameterArea = 0M;
            public bool IsManual = false;
            public SpecialToolsList SpecialToolForm = new SpecialToolsList();
            public string BasePath = System.IO.Path.GetDirectoryName(Application.ExecutablePath);

        #endregion Global Variables

        // Global Functions
        #region Global Functions
                // Function to convert Mixed fractions to decimals
            public decimal Convert_Decimal(string MixedFraction)
            {
                string[] SeparateDigits = (MixedFraction.Trim()).Split(' ');
                string[] SeparateFraction;
                decimal CompletePart = 0M;
                decimal Numerator = 0M;
                decimal Denominator = 0M;
                decimal ConvertedValue = 0M;

                if (SeparateDigits.Count() == 1)
                {
                    SeparateFraction = SeparateDigits[0].Split('/');
                    if (SeparateFraction.Count() > 1)
                    {
                        Numerator = Convert.ToInt32(SeparateFraction[0]);
                        Denominator = Convert.ToInt32(SeparateFraction[1]);
                    }
                    else
                    {
                        Numerator = Convert.ToInt32(SeparateFraction[0]);
                        Denominator = 1M;
                    }
                }
                else if (SeparateDigits.Count() > 1)
                {
                    SeparateFraction = SeparateDigits[1].Split('/');
                    CompletePart = Convert.ToInt32(SeparateDigits[0]);
                    if (SeparateFraction.Count() > 1)
                    {
                        Numerator = Convert.ToInt32(SeparateFraction[0]);
                        Denominator = Convert.ToInt32(SeparateFraction[1]);
                    }
                    else
                    {
                        MessageBox.Show("Bolt nominal diameter or bolt thread value is wrong. Cannot calculate the LTF.");
                        this.Close();
                    }
                }

                if (Denominator > 0)
                {
                    decimal Ratio = Numerator / Denominator;
                    ConvertedValue = CompletePart + Ratio;
                }
                return ConvertedValue;
            }

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
        #endregion  Global Functions
// End of global functions


// Main

        public Main()
        {
            InitializeComponent();
        }
        public Main(string ProjectFile)
        {
            InitializeComponent();
            OpenProjectFile(ProjectFile);
        }
        protected void AssignInitially()
        {
            SQLiteConnection conn = new SQLiteConnection("Data Source=" + BasePath + "\\WizBolt.db;New=False;");
            SQLiteCommand Populate_SQLiteCommand = new SQLiteCommand(conn);
            conn.Open();

            // Specifications
            try
            {
                Populate_SQLiteCommand.CommandText = "SELECT  (S.StandardId || '~' || SA.StandardAuxiliaryId) AS PrimaryStandardId, (S.StandardAbbreviation || ' ' || SA.StandardAuxiliary) AS Standard FROM Standards    S	INNER JOIN StandardAuxiliaries    SA	ON			SA.StandardId	=	S.StandardId   WHERE S.StandardId NOT IN (1, 2, 4)";
                DataSet Specifications_DataSet = new DataSet();
                SQLiteDataAdapter Specifications_DataAdapter = new SQLiteDataAdapter(Populate_SQLiteCommand);
                int SpecificationsResult = Specifications_DataAdapter.Fill(Specifications_DataSet);
                if (SpecificationsResult >= 0)
                {
                    Specification_ComboBox.DataSource = Specifications_DataSet.Tables[0];
                    Specification_ComboBox.DisplayMember = "Standard";
                    Specification_ComboBox.ValueMember = "PrimaryStandardId";
                    Specification_ComboBox.SelectedValue = "3~1";
                    Specification_ComboBox.SelectedItem = "ANSI B16.5";
                    SpecificationInTitle_Label.Text = ((DataRowView)Specification_ComboBox.SelectedItem)["Standard"].ToString();
                    SpecificationInfo_Label.Text = SpecificationInTitle_Label.Text;
                }
            }
            catch (Exception Specification_Exception)
            {
                MessageBox.Show(Specification_Exception.Message + " while populating Specifications data initially.");
            }

            // Flange ratings
            try                   // Flange Ratings
            {
                string SelectedSpecification = Specification_ComboBox.SelectedValue.ToString();
                string[] SelectedSpecificationSplit = SelectedSpecification.Split('~');
                Populate_SQLiteCommand.CommandText = "SELECT (SA.StandardId || '~' || SA.StandardAuxiliaryId || '~' || FR.FlangeRatingId) AS PrimaryFlangeRatingId, AbbreviatedRating FROM FlangeRatings FR "
                                                            + "INNER JOIN StandardAuxiliaries SA ON SA.StandardId = FR.StandardId  AND SA.StandardAuxiliaryId = FR.StandardAuxiliaryId WHERE SA.StandardId = "
                                                            + SelectedSpecificationSplit[0] + " AND SA.StandardAuxiliaryId = " + SelectedSpecificationSplit[1];
                DataSet FlangeRating_DataSet = new DataSet();
                SQLiteDataAdapter FlangeRating_DataAdapter = new SQLiteDataAdapter(Populate_SQLiteCommand);
                int FlangeRatingsResult = FlangeRating_DataAdapter.Fill(FlangeRating_DataSet);
                if ((FlangeRatingsResult > 0) && (FlangeRating_DataSet.Tables[0].Rows.Count > 0))
                {
                    Rating_ComboBox.DataSource = FlangeRating_DataSet.Tables[0];
                    Rating_ComboBox.DisplayMember = "AbbreviatedRating";
                    Rating_ComboBox.ValueMember = "PrimaryFlangeRatingId";
                    Rating_ComboBox.SelectedItem = "5.00 x 150 ANSI";
                    Rating_ComboBox.SelectedText = "5.00 x 150 ANSI";
                    Rating_ComboBox.SelectedValue = "3~1~10";
                    RatingInTitle_Label.Text = ((DataRowView)Rating_ComboBox.SelectedItem)["AbbreviatedRating"].ToString();
                    FlangeName_Label.Text = RatingInTitle_Label.Text;
                }
            }
            catch (Exception FlangeRating_Exception)
            {
                MessageBox.Show(FlangeRating_Exception.Message + " while populating flange ratings data initially.");
            }

            // Bolt Material
            try
            {
                Populate_SQLiteCommand.CommandText = " SELECT   M.MaterialId,    (S.StandardAbbreviation || ' ' || SA.StandardAuxiliary || '-' || COALESCE(M.MaterialAbbreviation, ' ')) AS  Material, M.Yield	"
                                                     + " FROM Materials 		M "
                                                     + " INNER JOIN StandardAuxiliaries  SA      ON          SA.StandardAuxiliaryId		=		M.StandardAuxiliaryId "
                                                     + " 											 AND     SA.StandardId 				=		M.StandardId "
                                                     + " INNER JOIN Standards   			S		ON			S.StandardId				=		M.StandardId "
                                                     + " 	WHERE 		SA.StandardId = 1	"
                                                     + " 				AND SA.StandardAuxiliaryId IN (1,  4); ";
                DataSet BoltMaterial_DataSet = new DataSet();
                SQLiteDataAdapter BoltMaterial_DataAdapter = new SQLiteDataAdapter(Populate_SQLiteCommand);
                int BoltMaterialResult = BoltMaterial_DataAdapter.Fill(BoltMaterial_DataSet);
                if ((BoltMaterialResult > 0) && (BoltMaterial_DataSet.Tables[0].Rows.Count > 0))
                {
                    BoltMaterial_ComboBox.DataSource = BoltMaterial_DataSet.Tables[0];
                    BoltMaterial_ComboBox.DisplayMember = "Material";
                    BoltMaterial_ComboBox.ValueMember = "MaterialId";
                    BoltMaterial_ComboBox.SelectedItem = "ASTM A193-B7";
                    BoltMaterial_ComboBox.SelectedText = "ASTM A193-B7";
                    BoltMaterial_ComboBox.SelectedValue = "3";

                    // Joint Info
                    BoltMaterial_Info_Label.Text = ((DataRowView)BoltMaterial_ComboBox.SelectedItem)["Material"].ToString();
                    BoltYield_SI_Label.Text = BoltMaterial_DataSet.Tables[0].Rows[2]["Yield"].ToString().Trim() + " MPa";
                    BoltYield_In_Label.Text = Math.Round((Convert.ToDecimal(BoltMaterial_DataSet.Tables[0].Rows[2]["Yield"].ToString()) * 0.145037738007M), 4, MidpointRounding.AwayFromZero).ToString().Trim() + " ksi";

                }
            }
            catch (Exception BoltMaterial_Exception)
            {
                MessageBox.Show(BoltMaterial_Exception.Message + " while populating bolt material data initially.");
            }

            // Flanges working side
            try
            {
                Populate_SQLiteCommand.CommandText = "SELECT * FROM FlangeTypes WHERE FlangeTypeId IN (2, 3, 111)";
                DataSet Flanges_DataSet = new DataSet();
                SQLiteDataAdapter Flanges_DataAdapter = new SQLiteDataAdapter(Populate_SQLiteCommand);
                int FlangesResult = Flanges_DataAdapter.Fill(Flanges_DataSet);
                if (FlangesResult >= 0)
                {

                    Flange1Config_ComboBox.DataSource = Flanges_DataSet.Tables[0];
                    Flange1Config_ComboBox.DisplayMember = "FlangeAbbreviation";
                    Flange1Config_ComboBox.ValueMember = "FlangeTypeId";
                    Flange1Config_ComboBox.SelectedItem = "WN-RF";
                    Flange1Config_ComboBox.SelectedText = "WN-RF";
                    Flange1Config_ComboBox.SelectedValue = 2;
                }

            }
            catch (Exception Flange_Exception)
            {
                MessageBox.Show(Flange_Exception.Message + " while populating flange data initially.");
            }

            // Flanges opposite
            try
            {
                Populate_SQLiteCommand.CommandText = "SELECT * FROM FlangeTypes WHERE FlangeTypeId IN (2, 3, 5, 6, 111)";
                DataSet FlangesOpposite_DataSet = new DataSet();
                SQLiteDataAdapter FlangesOpposite_DataAdapter = new SQLiteDataAdapter(Populate_SQLiteCommand);
                int FlangesOpposite_Result = FlangesOpposite_DataAdapter.Fill(FlangesOpposite_DataSet);
                if (FlangesOpposite_Result >= 0)
                {

                    Flange2Config_ComboBox.DataSource = FlangesOpposite_DataSet.Tables[0];
                    Flange2Config_ComboBox.DisplayMember = "FlangeAbbreviation";
                    Flange2Config_ComboBox.ValueMember = "FlangeTypeId";
                    Flange2Config_ComboBox.SelectedItem = "WN-RF";
                    Flange2Config_ComboBox.SelectedText = "WN-RF";
                    Flange2Config_ComboBox.SelectedValue = 2;
                }

            }
            catch (Exception Flange_Exception)
            {
                MessageBox.Show(Flange_Exception.Message + " while populating flange data initially.");
            }

            // Clamp Length, gasket, bolts, Residual Bolt Stress
            string FlangeDetails = Rating_ComboBox.SelectedValue.ToString();
            string[] FlangeId = FlangeDetails.Split('~');
            try
            {


                Populate_SQLiteCommand.CommandText = "SELECT    (FR.StandardId  ||  '~'  ||  FR.StandardAuxiliaryId  ||  '~'  ||  FR.FlangeRatingId  ||  '~'  ||  F.FlangeId)   AS PrimaryFlangeId, F.FlangeSize, F.NominalSize,  F.FlangeTypeId, F.FlangeThickness, F.SealRingGasketThickness, F.Number_of_Bolt_Holes, ResidualBoltStress   "
                                                  + "  FROM  Flanges		F	"
                                                  + "       INNER JOIN	FlangeRatings		FR			ON				FR.StandardId					=		F.StandardId    "
                                                  + "								AND		FR.StandardAuxiliaryId		=		F.StandardAuxiliaryId   "
                                                  + "								AND		FR.FlangeRatingId			=		F.FlangeRatingId    "
                                                  + " WHERE					F.FlangeTypeId			=	2   "
                                                  + "  AND		F.StandardId			=  " + FlangeId[0]
                                                  + "  AND		F.StandardAuxiliaryId  	=  " + FlangeId[1]
                                                  + "  AND  	F.FlangeRatingId 		=  " + FlangeId[2];

                DataSet Assembly_DataSet = new DataSet();
                SQLiteDataAdapter Assembly_DataAdapter = new SQLiteDataAdapter(Populate_SQLiteCommand);
                int AssemblyResult = Assembly_DataAdapter.Fill(Assembly_DataSet);
                if ((AssemblyResult > 0) && (Assembly_DataSet.Tables[0].Rows.Count > 0))
                {
                    Clamp1_TextBox.Text = (Math.Round((Convert.ToDecimal(Assembly_DataSet.Tables[0].Rows[0].ItemArray[4].ToString()) / 25.4M), 4, MidpointRounding.AwayFromZero)).ToString();
                    Clamp2_TextBox.Text = (Math.Round((Convert.ToDecimal(Assembly_DataSet.Tables[0].Rows[0].ItemArray[4].ToString()) / 25.4M), 4, MidpointRounding.AwayFromZero)).ToString();
                    GasketGap_TextBox.Text = (Math.Round((Convert.ToDecimal(Assembly_DataSet.Tables[0].Rows[0].ItemArray[5].ToString()) / 25.4M), 4, MidpointRounding.AwayFromZero)).ToString();
                    NumberOfBolts_TextBox.Text = Assembly_DataSet.Tables[0].Rows[0].ItemArray[6].ToString();
                    ResidualBoltStress_TextBox.Text = (Math.Round((Convert.ToDecimal(Assembly_DataSet.Tables[0].Rows[0].ItemArray[7].ToString()) * 145.037738007M), MidpointRounding.AwayFromZero)).ToString();
                    NoOfBoltsInfo_Label.Text = NumberOfBolts_TextBox.Text;
                }
            }
            catch (Exception Clamp_Exception)
            {
                MessageBox.Show(Clamp_Exception.Message + " while populating clamp data initially.");
            }

            // Bolt thread, pitch or threads per inch
            try
            {
                Populate_SQLiteCommand.CommandText = "SELECT    (F.StandardId  ||  '~'  ||  F.StandardAuxiliaryId  ||  '~'  ||  F.FlangeRatingId  ||  '~'  ||  F.FlangeId || '~' || FB.FlangeBoltId)   AS PrimaryFlangeBoltId, FB.BoltThread, FB.TPI_Pitch, FB.BoltLength   "
                                                      + "   FROM  FlangeBolts		FB  "
                                                      + "   INNER JOIN	Flanges					F			ON					F.StandardId					=		FB.StandardId   "
                                                      + "                               AND		F.StandardAuxiliaryId		=		FB.StandardAuxiliaryId  "
                                                      + "                               AND		F.FlangeRatingId				=		FB.FlangeRatingId   "
                                                      + "								AND		F.FlangeId						=		FB.FlangeId         "
                                                      + " WHERE					F.FlangeTypeId			= 2   "
                                                      + "               AND		F.StandardId			=    " + FlangeId[0]
                                                      + "               AND		F.StandardAuxiliaryId  	=    " + FlangeId[1]
                                                      + "               AND  	F.FlangeRatingId 		=    " + FlangeId[2];

                DataSet FlangeBolts_DataSet = new DataSet();
                SQLiteDataAdapter FlangeBolts_DataAdapter = new SQLiteDataAdapter(Populate_SQLiteCommand);
                int FlangeBoltsResult = FlangeBolts_DataAdapter.Fill(FlangeBolts_DataSet);
                if ((FlangeBoltsResult > 0) && (FlangeBolts_DataSet.Tables[0].Rows.Count > 0))
                {
                    BoltThread_ComboBox.DataSource = FlangeBolts_DataSet.Tables[0];
                    BoltThread_ComboBox.DisplayMember = "BoltThread";
                    BoltThread_ComboBox.ValueMember = "PrimaryFlangeBoltId";
 
                    PitchInfo_Label.Text = FlangeBolts_DataSet.Tables[0].Rows[0].ItemArray[2].ToString();
                    NominalThreadSizeInfo_Label.Text = ((DataRowView)BoltThread_ComboBox.SelectedItem)["BoltThread"].ToString() + "-" + PitchInfo_Label.Text + "UN";
                    Pitch_TextBox.Text = PitchInfo_Label.Text;
                }
            }
            catch (Exception FlangeBolts_Exception)
            {
                MessageBox.Show(FlangeBolts_Exception.Message + " while populating flange bolt data initially.");
            }

            string SelectedBoltThread = BoltThread_ComboBox.SelectedValue.ToString();
            string[] ToolParams = SelectedBoltThread.Split('~');
            string SelectedTool_List_of_Ids = string.Empty;
            string[] Tool_Identities = ToolSeriesId_Label.Text.Trim().Split('&');


            for (int i = 0; i < Tool_Identities.Count(); i++)
            {
                if (i == 0)
                {
                    SelectedTool_List_of_Ids = Tool_Identities[i];
                }
                else
                {
                    SelectedTool_List_of_Ids = SelectedTool_List_of_Ids + ", " + Tool_Identities[i];
                }
            }


            // Tensioner Tool
            //Populate_SQLiteCommand.CommandText = "SELECT (F.TensionerToolSeriesId || '~' || F.TensionerToolId) AS ToolId,  T.ModelNumber, T.HydraulicArea"
            //                                                            + "  FROM  Flanges		F	"
            //                                                            + "          INNER JOIN   TensionerTools  T    ON       T.TensionerToolSeriesId   =  F.TensionerToolSeriesId "
            //                                                            + "                                                 AND T.TensionerToolId         =  F.TensionerToolId  "
            //                                                             + " WHERE		F.FlangeTypeId			=	2"
            //                                                             + "                AND		F.StandardId			=  " + FlangeId[0]
            //                                                             + "                AND		F.StandardAuxiliaryId  	=  " + FlangeId[1]
            //                                                             + "                AND  	F.FlangeRatingId 		=  " + FlangeId[2];

            Populate_SQLiteCommand.CommandText = "   SELECT (F.TensionerToolSeriesId || '~' || F.TensionerToolId) AS ToolId,  T.ModelNumber, T.HydraulicArea "
                                                    + "  FROM  FlangeBoltTools		F	"
                                                    + "                 INNER JOIN   TensionerTools  T    ON       T.TensionerToolSeriesId   =  F.TensionerToolSeriesId "
                                                    + "                                                        AND T.TensionerToolId         =  F.TensionerToolId  "
                                                    + "                  WHERE		F.StandardId =   " + ToolParams[0]
                                                    + "                         AND F.StandardAuxiliaryId =  " + ToolParams[1]
                                                    + "                         AND F.FlangeRatingId   =  " + ToolParams[2]
                                                    + "                         AND F.FlangeId         =  " + ToolParams[3]
                                                    + "                         AND F.FlangeBoltId     =  " + ToolParams[4]
                                                    + "                         AND F.TensionerToolSeriesId IN (" + SelectedTool_List_of_Ids + ");";
            try
            {
                DataSet Tools_DataSet = new DataSet();
                SQLiteDataAdapter Tools_DataAdapter = new SQLiteDataAdapter(Populate_SQLiteCommand);
                int ToolResult = Tools_DataAdapter.Fill(Tools_DataSet);
                if ((ToolResult > 0) && (Tools_DataSet.Tables[0].Rows.Count > 0))
                {
                    TensionerTool_ComboBox.DataSource = Tools_DataSet.Tables[0];
                    TensionerTool_ComboBox.DisplayMember = "ModelNumber";
                    TensionerTool_ComboBox.ValueMember = "ToolId";
                    ToolId_Info_Label.Text = ((DataRowView)TensionerTool_ComboBox.SelectedItem)["ModelNumber"].ToString();
                    TensionerBolt_Info_Label.Text = NominalThreadSizeInfo_Label.Text;
                    string TPA = Tools_DataSet.Tables[0].Rows[0]["HydraulicArea"].ToString();
                    decimal ToolPressureArea = Math.Round((Convert.ToDecimal(TPA) / 645.16M), 4, MidpointRounding.AwayFromZero);
                    ToolPressureArea_Info_Label.Text = ToolPressureArea.ToString() + " Sq in";
                    ToolPressureArea_mm_Label.Text = TPA + " Sq mm";
                }
            }
            catch (Exception FlangeBolts_Exception)
            {
                MessageBox.Show(FlangeBolts_Exception.Message + " while populating flange bolt data initially.");
            }
            finally
            {
                conn.Close();
                conn.Dispose();
                Populate_SQLiteCommand.Dispose();
            }

            // Gaskets
            Gasket_ComboBox.Items.Clear();
            Gasket_ComboBox.Items.AddRange(new string[] {"", "Seal Ring", "RTJ", "Spiral", "Flat" });
            //Gasket_ComboBox.SelectedIndex = 1;
            Gasket_ComboBox.SelectedValue = "1";
            Gasket_ComboBox.SelectedItem = "Seal Ring";

            // Spacer
            Spacer_ComboBox.Items.Clear();
            Spacer_ComboBox.Items.AddRange(new string[] { "", "Single Washer", "Double Washers", "Auto Spacer" });

            // Bolt to tool proportion
            BoltTool_ComboBox.Items.Clear();
            BoltTool_ComboBox.Items.AddRange(new string[] { "100%", "50%", "25%", "TORQUE" });
            // BoltTool_ComboBox.SelectedIndex = 1;
            BoltTool_ComboBox.SelectedValue = "1";
            BoltTool_ComboBox.SelectedItem = "100%";
            IsInitial = false;

            // Graph
            DrawGraph();
        }
        private void Main_Load(object sender, EventArgs e)
        {
            AssignInitially();
        }

        private void Specification_ComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            string SelectedSpecification = Specification_ComboBox.SelectedValue.ToString();
            string[] SelectedSpecificationId = SelectedSpecification.Split('~');
            if (SelectedSpecification != "System.Data.DataRowView")
            {
                if ((Convert.ToInt32(SelectedSpecificationId[0]) == 0) && (Convert.ToInt32(SelectedSpecificationId[1]) == 0))
                {
                    IsManual = true;
                }
                else
                {
                    IsManual = false;
                    Clamp1_TextBox.BackColor = Color.White;
                    Clamp2_TextBox.BackColor = Color.White;
                    BoltMaterial_ComboBox.BackColor = Color.White;
                    BoltThread_ComboBox.BackColor = Color.White;
                    NumberOfBolts_TextBox.BackColor = Color.White;
                    NumberOfBolts_TextBox.ReadOnly = true;
                    ResidualBoltStress_TextBox.BackColor = Color.White;
                    BoltTool_ComboBox.Enabled = true;
                }

                SQLiteConnection connect = new SQLiteConnection("Data Source=" + BasePath + "\\WizBolt.db;New=False;");
                SQLiteCommand PopulateSpec_SQLiteCommand = new SQLiteCommand(connect);
                connect.Open();

                try
                {
                    PopulateSpec_SQLiteCommand.CommandText = "SELECT (SA.StandardId || '~' || SA.StandardAuxiliaryId || '~' || FR.FlangeRatingId) AS PrimaryFlangeRatingId, AbbreviatedRating FROM FlangeRatings FR "
                                                                + "INNER JOIN StandardAuxiliaries SA ON SA.StandardId = FR.StandardId  AND SA.StandardAuxiliaryId = FR.StandardAuxiliaryId WHERE SA.StandardId = "
                                                                + SelectedSpecificationId[0] + " AND SA.StandardAuxiliaryId = " + SelectedSpecificationId[1];

                    DataSet FlangeRating_DataSet = new DataSet();
                    SQLiteDataAdapter FlangeRating_DataAdapter = new SQLiteDataAdapter(PopulateSpec_SQLiteCommand);
                    int FlangeRatingsResult = FlangeRating_DataAdapter.Fill(FlangeRating_DataSet);
                    if ((FlangeRatingsResult > 0) && (FlangeRating_DataSet.Tables[0].Rows.Count > 0))
                    {
                        // Assign Ratings
                        Rating_ComboBox.DataSource = FlangeRating_DataSet.Tables[0];
                        Rating_ComboBox.DisplayMember = "AbbreviatedRating";
                        Rating_ComboBox.ValueMember = "PrimaryFlangeRatingId";

                        

                        // In Title 
                        SpecificationInTitle_Label.Text = ((DataRowView)Specification_ComboBox.SelectedItem)["Standard"].ToString();
                       // SpecificationInTitle_Label.Text = Specification_ComboBox.SelectedItem.ToString();
                        SpecificationInfo_Label.Text = SpecificationInTitle_Label.Text;
                    }
                    if ((IsManual) && (!IsLoading) && (!IsSetting))
                    {
                        PopulateSpec_SQLiteCommand.CommandText = "SELECT DISTINCT FT.FlangeTypeId, FT.FlangeAbbreviation"
                                                  + "  FROM  Flanges		F	"
                                                  + "       INNER JOIN	FlangeRatings		FR			ON				FR.StandardId					=		F.StandardId    "
                                                  + "								AND		FR.StandardAuxiliaryId		=		F.StandardAuxiliaryId   "
                                                  + "								AND		FR.FlangeRatingId			=		F.FlangeRatingId    "
                                                  + "       INNER JOIN FlangeTypes         FT           ON              FT.FlangeTypeId     =   F.FlangeTypeId"
                                                  + " WHERE		F.StandardId			=  0"
                                                  + "  AND		F.StandardAuxiliaryId  	=  0"
                                                  + "  AND  	F.FlangeRatingId 		=  1";
                                                  

                        DataSet FlangeTypes_DataSet = new DataSet();
                        SQLiteDataAdapter FlangeTypes_DataAdapter = new SQLiteDataAdapter(PopulateSpec_SQLiteCommand);
                        int FlangeTypeResult = FlangeTypes_DataAdapter.Fill(FlangeTypes_DataSet);
                        if (FlangeTypeResult > 0)
                        {
                            Flange1Config_ComboBox.DataSource = FlangeTypes_DataSet.Tables[0];
                            Flange1Config_ComboBox.DisplayMember = "FlangeAbbreviation";
                            Flange1Config_ComboBox.ValueMember = "FlangeTypeId";
                        }

                        PopulateSpec_SQLiteCommand.CommandText = "SELECT DISTINCT FT.FlangeTypeId, FT.FlangeAbbreviation"
                                                  + "  FROM  Flanges		F	"
                                                  + "       INNER JOIN	FlangeRatings		FR			ON				FR.StandardId					=		F.StandardId    "
                                                  + "								AND		FR.StandardAuxiliaryId		=		F.StandardAuxiliaryId   "
                                                  + "								AND		FR.FlangeRatingId			=		F.FlangeRatingId    "
                                                  + "       INNER JOIN FlangeTypes         FT           ON              FT.FlangeTypeId     =   F.FlangeTypeId"
                                                  + " WHERE		F.StandardId			= 0 " 
                                                  + "  AND		F.StandardAuxiliaryId  	= 0 " 
                                                  + "  AND  	F.FlangeRatingId 		= 1 " ;

                        DataSet FlangeOppositeTypes_DataSet = new DataSet();
                        SQLiteDataAdapter FlangeOppositeTypes_DataAdapter = new SQLiteDataAdapter(PopulateSpec_SQLiteCommand);
                        int FlangeOppositeTypeResult = FlangeOppositeTypes_DataAdapter.Fill(FlangeOppositeTypes_DataSet);
                        if (FlangeOppositeTypeResult > 0)
                        {
                            Flange2Config_ComboBox.DataSource = FlangeOppositeTypes_DataSet.Tables[0];
                            Flange2Config_ComboBox.DisplayMember = "FlangeAbbreviation";
                            Flange2Config_ComboBox.ValueMember = "FlangeTypeId";
                            Flange2Config_ComboBox.SelectedValue = FlangeOppositeTypes_DataSet.Tables[0].Rows[0]["FlangeTypeId"];
                            string SelectedBottomFlange = ((DataRowView)Flange2Config_ComboBox.SelectedItem)["FlangeAbbreviation"].ToString();
                            //Int64 SelectedBottomValue = Convert.ToInt64(((DataRowView)Flange2Config_ComboBox.SelectedValue)["FlangeTypeId"]);
                            // = SelectedBottomValue;
                            Flange2Config_ComboBox.SelectedItem = SelectedBottomFlange;
                        }
                        string SelectedFlange = Flange1Config_ComboBox.SelectedValue.ToString();
                        Clamp1_TextBox.Text = "";
                        Clamp1_TextBox.BackColor = Color.LightSalmon;
                        Clamp2_TextBox.Text = "";
                        Clamp2_TextBox.BackColor = Color.LightSalmon;
                        BoltMaterial_ComboBox.BackColor = Color.LightSalmon;
                        TensionerTool_ComboBox.DataSource = null;
                        TensionerTool_ComboBox.DisplayMember = null;
                        TensionerTool_ComboBox.ValueMember = null;
                        GasketGap_TextBox.Text = "0";
                        ResidualBoltStress_TextBox.Text = "";
                        ResidualBoltStress_TextBox.BackColor = Color.LightSalmon;
                        NumberOfBolts_TextBox.Text = "";
                        NumberOfBolts_TextBox.ReadOnly = false;
                        NumberOfBolts_TextBox.BackColor = Color.LightSalmon;
                        try
                        {
                            PopulateSpec_SQLiteCommand.CommandText = "SELECT (F.StandardId  ||  '~'  ||  F.StandardAuxiliaryId  ||  '~'  ||  F.FlangeRatingId  ||  '~'  ||  F.FlangeId || '~' || FB.FlangeBoltId) AS PrimaryFlangeBoltId, FB.BoltThread, FB.TPI_Pitch"
                                                                    + "  FROM  Flanges		F"
                                                                    + "        INNER JOIN	FlangeRatings		FR			ON				FR.StandardId					=		F.StandardId    "
                                                                    + "  							AND		FR.StandardAuxiliaryId		=		F.StandardAuxiliaryId   "
                                                                    + "  							AND		FR.FlangeRatingId			=		F.FlangeRatingId    "
                                                                    + "        INNER JOIN FlangeBolts          FB          ON              FB.StandardId       =       F.StandardId   "
                                                                    + "  							AND		FB.StandardAuxiliaryId		=		F.StandardAuxiliaryId   "
                                                                    + "  							AND		FB.FlangeRatingId			=		F.FlangeRatingId    "
                                                                    + "  							AND		FB.FlangeId					=		F.FlangeId         "
                                                                    + "  WHERE					F.FlangeTypeId			=	" + SelectedFlange
                                                                    + "   AND		F.StandardId			=  " + SelectedSpecificationId[0]
                                                                    + "   AND		F.StandardAuxiliaryId  	=  " + SelectedSpecificationId[1]
                                                                    + "   AND  	F.FlangeRatingId 		=  1;";

                             DataSet BoltSize_DataSet = new DataSet();
                             SQLiteDataAdapter BoltSize_DataAdapter = new SQLiteDataAdapter(PopulateSpec_SQLiteCommand);
                             int BoltSizeResult = BoltSize_DataAdapter.Fill(BoltSize_DataSet);
                             if (BoltSizeResult > 0)
                             {
                                 BoltThread_ComboBox.DataSource = BoltSize_DataSet.Tables[0];
                                 BoltThread_ComboBox.DisplayMember = "BoltThread";
                                 BoltThread_ComboBox.ValueMember = "PrimaryFlangeBoltId";
                                 Pitch_TextBox.Text = BoltSize_DataSet.Tables[0].Rows[0].ItemArray[2].ToString();
                                 Pitch_Label.Text = "T. P. I.";
                                 NominalThreadSizeInfo_Label.Text = ((DataRowView)BoltThread_ComboBox.SelectedItem)["BoltThread"].ToString().Trim() + "-" + Pitch_TextBox.Text.Trim() + "UN";
                                 BoltThread_ComboBox.BackColor = Color.LightSalmon;
                             }
                             BoltTool_ComboBox.SelectedValue = "3";
                             BoltTool_ComboBox.SelectedItem = "TORQUE";
                             BoltTool_ComboBox.Enabled = false;
                            // Title
                             RatingInTitle_Label.Text = ((DataRowView)Rating_ComboBox.SelectedItem)["AbbreviatedRating"].ToString();

                            // Info Panel
                             FlangeName_Label.Text = RatingInTitle_Label.Text;
                             TPI_Label.Text = "TPI";
                             PitchInfo_Label.Text = Pitch_TextBox.Text;
                             NoOfBoltsInfo_Label.Text = "";
                             BoltToolInfo_Label.Text = "";
                             ClampLengthInfo_Label.Text = "";
                             ClampLengthInfo_mm_Label.Text = "";
                             LTF_Label.Text = "";
                             ToolId_Info_Label.Text = "";
                             TensionerBolt_Info_Label.Text = "";
                             ToolPressureArea_Info_Label.Text = "";
                             ToolPressureArea_mm_Label.Text = "";
                             BoltMaterial_Info_Label.Text = ((DataRowView)BoltMaterial_ComboBox.SelectedItem)["Material"].ToString();
                    //        // BoltYield_SI_Label.Text = "";
                    //       //  BoltYield_In_Label.Text = "";
                             TensileStressArea_SI_Label.Text = "";
                             TensileStressArea_In_Label.Text = "";
                             BoltLength_In_Label.Text = "";
                             BoltLength_SI_Label.Text = "";
                             PressureA_In_Label.Text = "";
                             PressureA_SI_Label.Text = "";
                             PressureB_In_Label.Text = "";
                             PressureB_SI_Label.Text = "";
                             Detensioning_In_Label.Text = "";
                             Detensioning_SI_Label.Text = "";
                             Torque_In_Label.Text = "";
                             Torque_SI_Label.Text = "";
                             // CoefficientValue_Label.Text = "";

                             // Tabs
                             // Bolt Stress
                             // Row 1
                             T1ABoltStress_In_Label.Text = "";
                             T1ABoltStress_SI_Label.Text = "";
                             T1ABoltLoad_In_Label.Text = "";
                             T1ABoltLoad_SI_Label.Text = "";
                             T1ABoltYield_SI_Label.Text = "";

                             // Row 2
                             T1BBoltStress_In_Label.Text = "";
                             T1BBoltStress_SI_Label.Text = "";
                             T1BBoltLoad_In_Label.Text = "";
                             T1BBoltLoad_SI_Label.Text = "";
                             T1BBoltYield_Per_Label.Text = "";

                             // Row 3
                             T2RBoltStress_In_Label.Text = "";
                             T2RBoltStress_SI_Label.Text = "";
                             T2RBoltLoad_In_Label.Text = "";
                             T2RBoltLoad_SI_Label.Text = "";
                             T2RBoltYield_Per_Label.Text = "";

                             // Row 4
                             T3RBoltStress_In_Label.Text = "";
                             T3RBoltStress_SI_Label.Text = "";
                             T3RBoltLoad_In_Label.Text = "";
                             T3RBoltLoad_SI_Label.Text = "";
                             T3RBoltYield_Per_Label.Text = "";

                             // Row 5
                             DetenBoltStress_In_Label.Text = "";
                             DetenBoltStress_SI_Label.Text = "";
                             DetenBoltLoad_In_Label.Text = "";
                             DetenBoltLoad_SI_Label.Text = "";
                             DetenBoltYield_Per_Label.Text = "";

                             // Torque Tab
                             TorqueTab_SI_Label.Text = "";
                             TorqueValueTab_In_Label.Text = "";

                             // Graph Tab
                             DrawGraph();


                             // Bolt Tab
                             NominalBoltDiameterValueTab_Label.Text = "";
                             NumberOfBoltsValueTab_Label.Text = "";
                             // BoltYieldStrengthValueTab_In_Label.Text = "";
                             // BoltYieldStrengthValueTab_SI_Label.Text = "";
                             TensileStressAreaValueTab_In_Label.Text = "";
                             TensileStressAreaValueTab_SI_Label.Text = "";
                             BoltLengthValueTab_In_Label.Text = "";
                             BoltLengthValueTab_SI_Label.Text = "";

                             // Sequence Tab
                             Sequence_PictureBox.Image = null;
                             Pass1AppliedPressure_SequenceTabValue_In_Label.Text = "";
                             Pass1AppliedPressure_SequenceTabValue_SI_Label.Text = "";
                             Pass2Bolt_SequenceTabValue_Label.Text = "";
                             Pass2AppliedPressure_SequenceTabValue_SI_Label.Text = "";
                             Pass2AppliedPressure_SequenceTabValue_In_Label.Text = "";
                             Pass3Bolt_SequenceTabValue_Label.Text = "";
                             Pass3AppliedPressure_SequenceTabValue_SI_Label.Text = "";
                             Pass4Bolt_SequenceTabValue_Label.Text = "";
                             Pass4AppliedPressure_SequenceTabValue_SI_Label.Text = "";
                             Pass3AppliedPressure_SequenceTabValue_In_Label.Text = "";
                             Pass4AppliedPressure_SequenceTabValue_In_Label.Text = "";
                            
                             Torque_SI_Sequence_Label.Text = "";
                             Torque_In_Sequence_Label.Text = "";

                        }
                        catch (Exception Manual_Exception)
                        {
                            MessageBox.Show(Manual_Exception.Message + " during manual assignment of bolt size.");
                        }
                    }
                }
                catch (Exception Spec_Exception)
                {
                    MessageBox.Show(Spec_Exception.Message + " while changing specifications as per user choice.");
                }
                finally
                {
                    connect.Close();
                    connect.Dispose();
                    PopulateSpec_SQLiteCommand.Dispose();
                }
                
            }
        }

        private void Rating_ComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            
                if (BoltTool_ComboBox.SelectedItem != null)
                {
                    // Variables
                    
                    if ((!IsManual) && (!IsLoading) && (!IsSetting))
                    {
                        AssignRating();
                    }
                    if ((IsManual) && (IsLoading) && (IsSetting))
                    {
                        AssignRating();
                    }
                    if ((IsManual) && (!IsLoading) && (IsSetting))
                    {
                        AssignRating();
                    }
                    if ((!IsManual) && (!IsLoading) && (IsSetting))
                    {
                        AssignRating();
                    }
                    if ((!IsManual) && (IsLoading) && (IsSetting))
                    {
                        AssignRating();
                    }
                    if ((!IsManual) && (IsLoading) && (!IsSetting))
                    {
                        AssignRating();
                    }
                }
            
        }

        protected void AssignRating()
        {
            string SelectedTopFlangeType = string.Empty;
            string SelectedTopFlange = string.Empty;
            string Bolt_Diameter = string.Empty;
            string SelectedRating = Rating_ComboBox.SelectedValue.ToString();
            string[] SelectedRatingId = SelectedRating.Split('~');
            SQLiteConnection ConnectClamp = new SQLiteConnection("Data Source=" + BasePath + "\\WizBolt.db;New=False;");
            SQLiteCommand PopulateClamps_SQLiteCommand = new SQLiteCommand(ConnectClamp);
            ConnectClamp.Open();

            // When rating changes, the clamp length and bolts change.

            if (SelectedRating != "System.Data.DataRowView")
            {
                // Assign Flange's configuration.
                // Clear previous assignment of flanges
                Flange1Config_ComboBox.DataSource = null;
                Flange2Config_ComboBox.DataSource = null;
                Flange1Config_ComboBox.Items.Clear();
                Flange2Config_ComboBox.Items.Clear();

                PopulateClamps_SQLiteCommand.CommandText = "SELECT DISTINCT FT.FlangeTypeId, FT.FlangeAbbreviation"
                                          + "  FROM  Flanges		F	"
                                          + "       INNER JOIN	FlangeRatings		FR			ON				FR.StandardId					=		F.StandardId    "
                                          + "								AND		FR.StandardAuxiliaryId		=		F.StandardAuxiliaryId   "
                                          + "								AND		FR.FlangeRatingId			=		F.FlangeRatingId    "
                                          + "       INNER JOIN FlangeTypes         FT           ON              FT.FlangeTypeId     =   F.FlangeTypeId"
                                          + " WHERE		F.StandardId			=  " + SelectedRatingId[0]
                                          + "  AND		F.StandardAuxiliaryId  	=  " + SelectedRatingId[1]
                                          + "  AND  	F.FlangeRatingId 		=  " + SelectedRatingId[2]
                                          + "  AND      FT.FlangeTypeId  NOT IN (4, 5, 6, 11)";

                DataSet FlangeTypes_DataSet = new DataSet();
                SQLiteDataAdapter FlangeTypes_DataAdapter = new SQLiteDataAdapter(PopulateClamps_SQLiteCommand);
                int FlangeTypeResult = FlangeTypes_DataAdapter.Fill(FlangeTypes_DataSet);
                if (FlangeTypeResult > 0)
                {
                    Flange1Config_ComboBox.DataSource = FlangeTypes_DataSet.Tables[0];
                    Flange1Config_ComboBox.DisplayMember = "FlangeAbbreviation";
                    Flange1Config_ComboBox.ValueMember = "FlangeTypeId";
                   // Flange1Config_ComboBox.Items.Add("");     // Cannot be done when data source property used

                }

                PopulateClamps_SQLiteCommand.CommandText = "SELECT DISTINCT FT.FlangeTypeId, FT.FlangeAbbreviation"
                                          + "  FROM  Flanges		F	"
                                          + "       INNER JOIN	FlangeRatings		FR			ON				FR.StandardId					=		F.StandardId    "
                                          + "								AND		FR.StandardAuxiliaryId		=		F.StandardAuxiliaryId   "
                                          + "								AND		FR.FlangeRatingId			=		F.FlangeRatingId    "
                                          + "       INNER JOIN FlangeTypes         FT           ON              FT.FlangeTypeId     =   F.FlangeTypeId"
                                          + " WHERE		F.StandardId			=  " + SelectedRatingId[0]
                                          + "  AND		F.StandardAuxiliaryId  	=  " + SelectedRatingId[1]
                                          + "  AND  	F.FlangeRatingId 		=  " + SelectedRatingId[2];

                DataSet FlangeOppositeTypes_DataSet = new DataSet();
                SQLiteDataAdapter FlangeOppositeTypes_DataAdapter = new SQLiteDataAdapter(PopulateClamps_SQLiteCommand);
                int FlangeOppositeTypeResult = FlangeOppositeTypes_DataAdapter.Fill(FlangeOppositeTypes_DataSet);
                if (FlangeOppositeTypeResult > 0)
                {
                    Flange2Config_ComboBox.DataSource = FlangeOppositeTypes_DataSet.Tables[0];
                    Flange2Config_ComboBox.DisplayMember = "FlangeAbbreviation";
                    Flange2Config_ComboBox.ValueMember = "FlangeTypeId";
                    // Flange2Config_ComboBox.Items.Add("");         // Cannot be done when data source property used
                    Flange2Config_ComboBox.SelectedValue = FlangeOppositeTypes_DataSet.Tables[0].Rows[0]["FlangeTypeId"];
                    string SelectedBottomFlange = ((DataRowView)Flange2Config_ComboBox.SelectedItem)["FlangeAbbreviation"].ToString();

                    Flange2Config_ComboBox.SelectedItem = SelectedBottomFlange;
                }

                SelectedTopFlangeType = Flange1Config_ComboBox.SelectedValue.ToString();
                SelectedTopFlange = Flange1Config_ComboBox.SelectedItem.ToString();

                try
                {
                    PopulateClamps_SQLiteCommand.CommandText = "SELECT    (FR.StandardId  ||  '~'  ||  FR.StandardAuxiliaryId  ||  '~'  ||  FR.FlangeRatingId  ||  '~'  ||  F.FlangeId)   AS PrimaryFlangeId, "
                                                      + " F.FlangeSize, F.NominalSize,  F.FlangeTypeId, F.FlangeThickness, F.SealRingGasketThickness, F.Number_of_Bolt_Holes, F.ResidualBoltStress,   "
                                                      + " (F.StandardId  ||  '~'  ||  F.StandardAuxiliaryId  ||  '~'  ||  F.FlangeRatingId  ||  '~'  ||  F.FlangeId || '~' || FB.FlangeBoltId) AS PrimaryFlangeBoltId, FB.BoltThread, FB.TPI_Pitch, FB.BoltLength "
                                                      + "  FROM  Flanges		F	"
                                                      + "       INNER JOIN	FlangeRatings		FR			ON				FR.StandardId					=		F.StandardId    "
                                                      + "								AND		FR.StandardAuxiliaryId		=		F.StandardAuxiliaryId   "
                                                      + "								AND		FR.FlangeRatingId			=		F.FlangeRatingId    "
                                                      + "       INNER JOIN FlangeBolts          FB          ON              FB.StandardId       =       F.StandardId   "
                                                      + "								AND		FB.StandardAuxiliaryId		=		F.StandardAuxiliaryId   "
                                                      + "								AND		FB.FlangeRatingId			=		F.FlangeRatingId    "
                                                      + "								AND		FB.FlangeId					=		F.FlangeId         "
                                                      + " WHERE					F.FlangeTypeId			=	" + SelectedTopFlangeType
                                                      + "  AND		F.StandardId			=  " + SelectedRatingId[0]
                                                      + "  AND		F.StandardAuxiliaryId  	=  " + SelectedRatingId[1]
                                                      + "  AND  	F.FlangeRatingId 		=  " + SelectedRatingId[2];

                    DataSet ClampedBolts_DataSet = new DataSet();
                    SQLiteDataAdapter ClampBolt_DataAdapter = new SQLiteDataAdapter(PopulateClamps_SQLiteCommand);
                    int ClampResult = ClampBolt_DataAdapter.Fill(ClampedBolts_DataSet);
                    if ((ClampResult > 0) && (ClampedBolts_DataSet.Tables[0].Rows.Count > 0))
                    {
                        if (UnitSystem_Button.Text == "SAE")
                        {
                            if ((ClampedBolts_DataSet.Tables[0].Rows[0].ItemArray[4].ToString().Trim() != "") && (ClampedBolts_DataSet.Tables[0].Rows[0].ItemArray[4].ToString().Trim().Length > 0))
                            {
                                Clamp1_TextBox.Text = (Math.Round((Convert.ToDecimal(ClampedBolts_DataSet.Tables[0].Rows[0].ItemArray[4].ToString()) / 25.4M), 4, MidpointRounding.AwayFromZero)).ToString();
                                Clamp2_TextBox.Text = (Math.Round((Convert.ToDecimal(ClampedBolts_DataSet.Tables[0].Rows[0].ItemArray[4].ToString()) / 25.4M), 4, MidpointRounding.AwayFromZero)).ToString();
                            }
                            if ((ClampedBolts_DataSet.Tables[0].Rows[0].ItemArray[5].ToString().Trim() != "") && (ClampedBolts_DataSet.Tables[0].Rows[0].ItemArray[5].ToString().Trim().Length > 0))
                            {
                                GasketGap_TextBox.Text = (Math.Round((Convert.ToDecimal(ClampedBolts_DataSet.Tables[0].Rows[0].ItemArray[5].ToString()) / 25.4M), 4, MidpointRounding.AwayFromZero)).ToString();
                            }
                            if ((ClampedBolts_DataSet.Tables[0].Rows[0].ItemArray[7].ToString().Trim() != "") && (ClampedBolts_DataSet.Tables[0].Rows[0].ItemArray[7].ToString().Trim().Length > 0))
                            {
                                ResidualBoltStress_TextBox.Text = (Math.Round((Convert.ToDecimal(ClampedBolts_DataSet.Tables[0].Rows[0].ItemArray[7].ToString()) * 145.037738007M), MidpointRounding.AwayFromZero)).ToString();
                            }
                        }
                        else
                        {
                            if ((ClampedBolts_DataSet.Tables[0].Rows[0].ItemArray[4].ToString().Trim() != "") && (ClampedBolts_DataSet.Tables[0].Rows[0].ItemArray[4].ToString().Trim().Length > 0))
                            {
                                Clamp1_TextBox.Text = Convert.ToDecimal(ClampedBolts_DataSet.Tables[0].Rows[0].ItemArray[4].ToString()).ToString();
                                Clamp2_TextBox.Text = Convert.ToDecimal(ClampedBolts_DataSet.Tables[0].Rows[0].ItemArray[4].ToString()).ToString();
                            }
                            if ((ClampedBolts_DataSet.Tables[0].Rows[0].ItemArray[5].ToString().Trim() != "") && (ClampedBolts_DataSet.Tables[0].Rows[0].ItemArray[5].ToString().Trim().Length > 0))
                            {
                                GasketGap_TextBox.Text = Convert.ToDecimal(ClampedBolts_DataSet.Tables[0].Rows[0].ItemArray[5].ToString()).ToString();
                            }
                            if ((ClampedBolts_DataSet.Tables[0].Rows[0].ItemArray[7].ToString().Trim() != "") && (ClampedBolts_DataSet.Tables[0].Rows[0].ItemArray[7].ToString().Trim().Length > 0))
                            {
                                ResidualBoltStress_TextBox.Text = Convert.ToDecimal(ClampedBolts_DataSet.Tables[0].Rows[0].ItemArray[7].ToString()).ToString();
                            }
                        }
                        // Spacer
                        Spacer_ComboBox.SelectedValue = "1";
                        Spacer_ComboBox.SelectedItem = "";
                        CustomSpacer_TextBox.Text = "";

                        // Bolts
                        NumberOfBolts_TextBox.Text = ClampedBolts_DataSet.Tables[0].Rows[0].ItemArray[6].ToString();

                        // Bolt Size drop down
                        BoltThread_ComboBox.DataSource = ClampedBolts_DataSet.Tables[0];
                        BoltThread_ComboBox.DisplayMember = "BoltThread";
                        BoltThread_ComboBox.ValueMember = "PrimaryFlangeBoltId";

                    }
                    // Title and Joint Information
                    RatingInTitle_Label.Text = ((DataRowView)Rating_ComboBox.SelectedItem)["AbbreviatedRating"].ToString();
                    //RatingInTitle_Label.Text = Rating_ComboBox.SelectedItem.ToString();
                    FlangeName_Label.Text = RatingInTitle_Label.Text;
                }

                catch (Exception Rating_Exception)
                {
                    MessageBox.Show(Rating_Exception.Message + " during rating change.");
                }
                finally
                {
                    ConnectClamp.Close();
                    PopulateClamps_SQLiteCommand.Dispose();
                    ConnectClamp.Dispose();
                }
            }
        }
        
        private void Flange1Config_ComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            //if (IsManual)
            //{
            //    Clamp1_TextBox.Text = "";
            //    Clamp1_TextBox.BackColor = Color.Red;
            //}
            //else
            //{
                if ((Flange1Config_ComboBox.SelectedValue != null) && (Rating_ComboBox.SelectedValue != null))
                {
                    string SelectedFlange = Flange1Config_ComboBox.SelectedValue.ToString();
                    string SelectedRatings = Rating_ComboBox.SelectedValue.ToString();
                    string[] SelectedRatingId = SelectedRatings.Split('~');


                    if ((SelectedFlange != "System.Data.DataRowView") && (SelectedRatings != "System.Data.DataRowView") && (!IsInitial))
                    {
                        SQLiteConnection FlangeConnect = new SQLiteConnection("Data Source=" + BasePath + "\\WizBolt.db;New=False;");
                        SQLiteCommand PopulateFlanges_SQLiteCommand = new SQLiteCommand(FlangeConnect);
                        FlangeConnect.Open();

                        try
                        {
                            PopulateFlanges_SQLiteCommand.CommandText = "SELECT    (FR.StandardId  ||  '~'  ||  FR.StandardAuxiliaryId  ||  '~'  ||  FR.FlangeRatingId  ||  '~'  ||  F.FlangeId)   AS PrimaryFlangeId,  F.FlangeThickness   "
                                                          + "  FROM  Flanges		F	"
                                                          + "       INNER JOIN	FlangeRatings		FR			ON				FR.StandardId					=		F.StandardId    "
                                                          + "								AND		FR.StandardAuxiliaryId		=		F.StandardAuxiliaryId   "
                                                          + "								AND		FR.FlangeRatingId			=		F.FlangeRatingId    "
                                                          + " WHERE		FlangeTypeId            =  " + SelectedFlange
                                                          + "  AND      F.StandardId			=  " + SelectedRatingId[0]
                                                          + "  AND		F.StandardAuxiliaryId  	=  " + SelectedRatingId[1]
                                                          + "  AND  	F.FlangeRatingId 		=  " + SelectedRatingId[2];


                            DataSet FlangeThickness_DataSet = new DataSet();
                            SQLiteDataAdapter FlangeThickness_DataAdapter = new SQLiteDataAdapter(PopulateFlanges_SQLiteCommand);
                            int FlangeResult = FlangeThickness_DataAdapter.Fill(FlangeThickness_DataSet);
                            if ((FlangeResult > 0) && (FlangeThickness_DataSet.Tables[0].Rows.Count > 0))
                            {
                                if ((FlangeThickness_DataSet.Tables[0].Rows[0].ItemArray[1].ToString().Trim() != "") && (FlangeThickness_DataSet.Tables[0].Rows[0].ItemArray[1].ToString().Trim().Length > 0))
                                {
                                    if (UnitSystem_Button.Text == "SAE")
                                    {
                                        Clamp1_TextBox.Text = (Math.Round((Convert.ToDecimal(FlangeThickness_DataSet.Tables[0].Rows[0].ItemArray[1].ToString()) / 25.4M), 4, MidpointRounding.AwayFromZero)).ToString();
                                    }
                                    else
                                    {
                                        Clamp1_TextBox.Text = FlangeThickness_DataSet.Tables[0].Rows[0].ItemArray[1].ToString();
                                    }
                                }
                                else
                                {
                                    Clamp1_TextBox.Text = "";

                                }
                            }
                        }
                        catch (Exception Flange_Exception)
                        {
                            MessageBox.Show(Flange_Exception.Message + " while changing flanges as per user choice.");
                        }
                        finally
                        {
                            FlangeConnect.Close();
                            FlangeConnect.Dispose();
                            PopulateFlanges_SQLiteCommand.Dispose();
                        }
                    }
                }
            //}
        }

        private void Flange2Config_ComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            //if (IsManual)
            //{
            //    Clamp2_TextBox.Text = "";
            //    Clamp2_TextBox.BackColor = Color.Red;
            //}
            //else
            //{
                if ((Flange2Config_ComboBox.SelectedValue != null) && (Rating_ComboBox.SelectedValue != null))
                {
                    string SelectedFlange = Flange2Config_ComboBox.SelectedValue.ToString();
                    string SelectedRatings = Rating_ComboBox.SelectedValue.ToString();
                    string[] SelectedRatingId = SelectedRatings.Split('~');

                    if ((SelectedFlange != "System.Data.DataRowView") && (SelectedRatings != "System.Data.DataRowView") && (!IsInitial))
                    {
                        SQLiteConnection FlangeConnect = new SQLiteConnection("Data Source=" + BasePath + "\\WizBolt.db;New=False;");
                        SQLiteCommand PopulateFlanges_SQLiteCommand = new SQLiteCommand(FlangeConnect);
                        FlangeConnect.Open();

                        try
                        {
                            PopulateFlanges_SQLiteCommand.CommandText = "SELECT    (FR.StandardId  ||  '~'  ||  FR.StandardAuxiliaryId  ||  '~'  ||  FR.FlangeRatingId  ||  '~'  ||  F.FlangeId)   AS PrimaryFlangeId,  F.FlangeThickness   "
                                                          + "  FROM  Flanges		F	"
                                                          + "       INNER JOIN	FlangeRatings		FR			ON				FR.StandardId					=		F.StandardId    "
                                                          + "								AND		FR.StandardAuxiliaryId		=		F.StandardAuxiliaryId   "
                                                          + "								AND		FR.FlangeRatingId			=		F.FlangeRatingId    "
                                                          + " WHERE		FlangeTypeId            =  " + SelectedFlange
                                                          + "  AND      F.StandardId			=  " + SelectedRatingId[0]
                                                          + "  AND		F.StandardAuxiliaryId  	=  " + SelectedRatingId[1]
                                                          + "  AND  	F.FlangeRatingId 		=  " + SelectedRatingId[2];


                            DataSet FlangeThickness_DataSet = new DataSet();
                            SQLiteDataAdapter FlangeThickness_DataAdapter = new SQLiteDataAdapter(PopulateFlanges_SQLiteCommand);
                            int FlangeResult = FlangeThickness_DataAdapter.Fill(FlangeThickness_DataSet);
                            if ((FlangeResult > 0) && (FlangeThickness_DataSet.Tables[0].Rows.Count > 0))
                            {
                                if ((FlangeThickness_DataSet.Tables[0].Rows[0].ItemArray[1].ToString().Trim() != "") && (FlangeThickness_DataSet.Tables[0].Rows[0].ItemArray[1].ToString().Trim().Length > 0))
                                {
                                    if (UnitSystem_Button.Text == "SAE")
                                    {
                                        Clamp2_TextBox.Text = (Math.Round((Convert.ToDecimal(FlangeThickness_DataSet.Tables[0].Rows[0].ItemArray[1].ToString()) / 25.4M), 4, MidpointRounding.AwayFromZero)).ToString();
                                    }
                                    else
                                    {
                                        Clamp2_TextBox.Text = FlangeThickness_DataSet.Tables[0].Rows[0].ItemArray[1].ToString();
                                    }
                                }
                                else
                                {
                                    Clamp2_TextBox.Text = "";
                                }
                            }
                        }
                        catch (Exception Flange_Exception)
                        {
                            MessageBox.Show(Flange_Exception.Message + " while changing flanges as per user choice.");
                        }
                        finally
                        {
                            FlangeConnect.Close();
                            FlangeConnect.Dispose();
                            PopulateFlanges_SQLiteCommand.Dispose();
                        }
                    }
                }
           // }
        }

        private void BoltThread_ComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (!IsInitial)
            {
                //if (!IsSetting)
                //{
                    // Local Variables
                    
                    string SelectedBoltThread = string.Empty;
                    int SelectedBoltIndex = 0;
                    int PreviousPitch = 0;					// BugFix: Torque and other parameters changed only when Pitch changed even though it depends on bolt diameter
                    int CurrentPitch = 0;					// BugFix: Torque and other parameters changed only when Pitch changed even though it depends on bolt diameter
                    object PitchSender = new object();
                    EventArgs PitchEvent = new EventArgs();
                    //if (!IsLoading)
                    //{
                    SQLiteConnection Bolt_Connection = new SQLiteConnection("Data Source=" + BasePath + "\\WizBolt.db;New=False;");
                        SQLiteCommand Bolt_Command = new SQLiteCommand(Bolt_Connection);
                        
                            IsBoltThread = true;
                            if (BoltThread_ComboBox.SelectedValue != null)
                            {
                                BoltThread_ComboBox.BackColor = Color.White;

                                // Set the Tensioner Tool in TensionerTool_ComboBox
                                SelectedBoltThread = BoltThread_ComboBox.SelectedValue.ToString();
                                string SelectedTool_List_of_Ids = string.Empty;
                                string[] Tool_Identities = ToolSeriesId_Label.Text.Trim().Split('&');
                                string[] ToolParams = SelectedBoltThread.Split('~');
                                string Auxiliary_and_Standard_Id = (Specification_ComboBox.SelectedValue).ToString();  // Standard Id and Auxiliary Id in the form of 1~1
                                string[] Separated_Id = Auxiliary_and_Standard_Id.Split('~');               // Standard Id and Auxiliary Ids split in individual

                                for (int i = 0; i < Tool_Identities.Count(); i++ )
                                {
                                    if (i == 0)
                                    {
                                        SelectedTool_List_of_Ids = Tool_Identities[i];
                                    }
                                    else
                                    {
                                        SelectedTool_List_of_Ids = SelectedTool_List_of_Ids + ", " + Tool_Identities[i];
                                    }
                                }

                                    if (Convert.ToInt32(Separated_Id[0]) == 0)                                  // Manual. No standard selected, but wants only maanual operation which is equivalent of customized. 
                                    {
                                        Bolt_Command.CommandText = "   SELECT DISTINCT (F.TensionerToolSeriesId || '~' || F.TensionerToolId) AS ToolId,  T.ModelNumber, T.HydraulicArea "
                                                                + "  FROM  FlangeBoltTools		F	"
                                                                + "                 INNER JOIN   TensionerTools  T    ON       T.TensionerToolSeriesId   =  F.TensionerToolSeriesId "
                                                                + "                                                        AND T.TensionerToolId         =  F.TensionerToolId  "
                                                                + "                  WHERE		F.StandardId =   " + ToolParams[0]
                                                                + "                         AND F.StandardAuxiliaryId =  " + ToolParams[1]
                                                                + "                         AND F.FlangeRatingId   =  " + ToolParams[2]
                                                                + "                         AND F.FlangeId         =  " + ToolParams[3]
                                                                + "                         AND F.FlangeBoltId     =  " + ToolParams[4]
                                                                + "                         AND F.TensionerToolSeriesId IN (5, " + SelectedTool_List_of_Ids +");";
                                    }
                                    else
                                    {
                                        Bolt_Command.CommandText = "   SELECT (F.TensionerToolSeriesId || '~' || F.TensionerToolId) AS ToolId,  T.ModelNumber, T.HydraulicArea "
                                                        + "  FROM  FlangeBoltTools		F	"
                                                        + "                 INNER JOIN   TensionerTools  T    ON       T.TensionerToolSeriesId   =  F.TensionerToolSeriesId "
                                                        + "                                                        AND T.TensionerToolId         =  F.TensionerToolId  "
                                                        + "                  WHERE		F.StandardId =   " + ToolParams[0]
                                                        + "                         AND F.StandardAuxiliaryId =  " + ToolParams[1]
                                                        + "                         AND F.FlangeRatingId   =  " + ToolParams[2]
                                                        + "                         AND F.FlangeId         =  " + ToolParams[3]
                                                        + "                         AND F.FlangeBoltId     =  " + ToolParams[4]
                                                        + "                         AND F.TensionerToolSeriesId IN (" + SelectedTool_List_of_Ids + ");";

                                    }

                                try
                                {
                                    DataSet Tools_DataSet = new DataSet();
                                    SQLiteDataAdapter Tools_DataAdapter = new SQLiteDataAdapter(Bolt_Command);
                                    int ToolResult = Tools_DataAdapter.Fill(Tools_DataSet);
                                    if ((ToolResult > 0) && (Tools_DataSet.Tables[0].Rows.Count > 0))
                                    {
                                        TensionerTool_ComboBox.DataSource = Tools_DataSet.Tables[0];
                                        TensionerTool_ComboBox.DisplayMember = "ModelNumber";
                                        TensionerTool_ComboBox.ValueMember = "ToolId";
                                    }
                                    else
                                    {
                                        TensionerTool_ComboBox.DataSource = null;
                                        TensionerTool_ComboBox.DisplayMember = null;
                                        TensionerTool_ComboBox.ValueMember = null;
                                    }
                                }
                                catch (Exception Tools_Exception)
                                {
                                    MessageBox.Show(Tools_Exception.Message + " while populating tools during bolt sizing event.");
                                }

                                SelectedBoltIndex = BoltThread_ComboBox.SelectedIndex;
                                DataTable PitchTable = (DataTable)BoltThread_ComboBox.DataSource;
                                Pitch_TextBox.Text = PitchTable.Rows[SelectedBoltIndex]["TPI_Pitch"].ToString();
                                // BugFix: Torque and other parameters changed only when Pitch changed even though it depends on bolt diameter
                                if (Pitch_TextBox.Text.Trim().Length > 0)
                                {
                                    try
                                    {
                                        PreviousPitch = Convert.ToInt32(Pitch_TextBox.Text.Trim());
                                    }
                                    catch
                                    {

                                    }
                                }
                                Pitch_TextBox.Text = PitchTable.Rows[SelectedBoltIndex]["TPI_Pitch"].ToString();
                                if (Pitch_TextBox.Text.Trim().Length > 0)
                                {
                                    try
                                    {
                                        CurrentPitch = Convert.ToInt32(Pitch_TextBox.Text.Trim());
                                    }
                                    catch
                                    {

                                    }
                                }
                                if (PreviousPitch == CurrentPitch)
                                {
                                    PitchSender = Pitch_TextBox;

                                    Pitch_TextBox_TextChanged(PitchSender, PitchEvent);
                                }
                                // BugFix: End
                                IsBoltThread = false;
                            }
                        
                   // }
               // }
            }
        }

        private void Spacer_ComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (!IsLoading)
            {
                if (!IsSetting)
                {
                    string SelectedSpacer = Spacer_ComboBox.SelectedItem.ToString();
                    if (SelectedSpacer == "Auto Spacer")
                    {
                        // Initialize all variables
                        decimal BoltDiameter = 0M;
                        decimal TopClamp = 0M;
                        decimal BottomClamp = 0M;
                        decimal GasketThickness = 0M;
                        decimal Min_LTF = 1.1M;
                        decimal Max_LTF = 0M;
                        decimal Spacer = 0M;
                        decimal CurrentClampLength = 0M;
                        decimal Entered_LTF = 0M;
                        // Check units
                        string BoltThread = ((DataRowView)BoltThread_ComboBox.SelectedItem)["BoltThread"].ToString();
                        //string BoltThread = BoltThread_ComboBox.SelectedItem.ToString();
                        //"System.Data.DataRowView"
                        if (BoltThread.Substring(0, 1) == "M")                          // Nominal Diameter is in mm
                        {
                            BoltDiameter = Convert.ToDecimal(BoltThread.Substring(1));      // Bolt size is in mm
                            if (UnitSystem_Button.Text == "SAE")                               // Rest of the units such as clamp length are in inches
                            {
                                if (Clamp1_TextBox.Text.Length > 0)
                                {
                                    TopClamp = Math.Round((Convert.ToDecimal(Clamp1_TextBox.Text) * 25.4M), 3, MidpointRounding.AwayFromZero);      // Convert to mm from inch
                                }

                                if (Clamp2_TextBox.Text.Length > 0)
                                {
                                    BottomClamp = Math.Round((Convert.ToDecimal(Clamp2_TextBox.Text) * 25.4M), 3, MidpointRounding.AwayFromZero);   // Convert to mm from inch
                                }
                                if (GasketGap_TextBox.Text.Length > 0)
                                {
                                    GasketThickness = Math.Round((Convert.ToDecimal(GasketGap_TextBox.Text) * 25.4M), 3, MidpointRounding.AwayFromZero);    // Convert to mm from inch
                                }
                            }
                            else        // Since Bolt size is Mxx it is in mm. In this case rest of the units such as clamp length are also in mm.
                            {
                                if (Clamp1_TextBox.Text.Length > 0)
                                {
                                    TopClamp = Convert.ToDecimal(Clamp1_TextBox.Text);
                                }
                                if (Clamp2_TextBox.Text.Length > 0)
                                {
                                    BottomClamp = Convert.ToDecimal(Clamp2_TextBox.Text);
                                }
                                if (GasketGap_TextBox.Text.Length > 0)
                                {
                                    GasketThickness = Convert.ToDecimal(GasketGap_TextBox.Text);
                                }
                            }

                        }
                        else               // Nominal diameter is in inch
                        {
                            BoltThread = BoltThread.Substring(0, BoltThread.Length - 1);
                            BoltDiameter = Convert_Decimal(BoltThread);
                            BoltDiameter = BoltDiameter * 25.4M;

                            if (UnitSystem_Button.Text == "SAE")                               // Rest of the units such as clamp length are in inches
                            {
                                if (Clamp1_TextBox.Text.Length > 0)
                                {
                                    TopClamp = Math.Round((Convert.ToDecimal(Clamp1_TextBox.Text) * 25.4M), 3, MidpointRounding.AwayFromZero);      // Convert to mm from inch
                                }

                                if (Clamp2_TextBox.Text.Length > 0)
                                {
                                    BottomClamp = Math.Round((Convert.ToDecimal(Clamp2_TextBox.Text) * 25.4M), 3, MidpointRounding.AwayFromZero);   // Convert to mm from inch
                                }
                                if (GasketGap_TextBox.Text.Length > 0)
                                {
                                    GasketThickness = Math.Round((Convert.ToDecimal(GasketGap_TextBox.Text) * 25.4M), 3, MidpointRounding.AwayFromZero);    // Convert to mm from inch
                                }
                            }
                            else        // Rest of the units such as clamp length are in mm
                            {
                                if (Clamp1_TextBox.Text.Length > 0)
                                {
                                    TopClamp = Convert.ToDecimal(Clamp1_TextBox.Text);
                                }
                                if (Clamp2_TextBox.Text.Length > 0)
                                {
                                    BottomClamp = Convert.ToDecimal(Clamp2_TextBox.Text);
                                }
                                if (GasketGap_TextBox.Text.Length > 0)
                                {
                                    GasketThickness = Convert.ToDecimal(GasketGap_TextBox.Text);
                                }
                            }
                        }
                        // All units are in mm
                        CurrentClampLength = TopClamp + BottomClamp + GasketThickness;
                        Max_LTF = 1.01M + Math.Round((BoltDiameter / CurrentClampLength), 4, MidpointRounding.AwayFromZero);

                        LTF LTF_Popup = new LTF();
                        LTF_Popup.Minimum_LTF = Min_LTF;
                        LTF_Popup.Maximum_LTF = Max_LTF;
                        LTF_Popup.ShowDialog();
                        Entered_LTF = LTF_Popup.Entered_LTF;
                        if (Entered_LTF > 1.1M)
                        {
                            Spacer = (BoltDiameter / (Entered_LTF - 1.01M)) - CurrentClampLength;
                            if (UnitSystem_Button.Text == "SAE")
                            {
                                Spacer = Math.Round((Spacer / 25.4M), 4, MidpointRounding.AwayFromZero);
                                CustomSpacer_TextBox.Text = Spacer.ToString();
                            }
                            else
                            {
                                CustomSpacer_TextBox.Text = Math.Round(Spacer, 0, MidpointRounding.AwayFromZero).ToString();
                            }
                            LTF_Label.Text = Entered_LTF.ToString();
                        }
                    }
                }
            }
        }

        protected decimal Calculate_LTF()
        {
            decimal CurrentLTF = 0M;

            // Initialize all variables
            decimal BoltDiameter = 0M;
            decimal TopClamp = 0M;
            decimal BottomClamp = 0M;
            decimal GasketThickness = 0M;
            decimal SpacerThickness = 0M;
            decimal CurrentClampLength = 0M;
            
            // Check units
            string BoltThread = ((DataRowView)BoltThread_ComboBox.SelectedItem)["BoltThread"].ToString();
            //string BoltThread = BoltThread_ComboBox.SelectedItem.ToString();
            //"System.Data.DataRowView"
            if (BoltThread.Substring(0, 1) == "M")                          // Nominal Diameter is in mm
            {
                BoltDiameter = Convert.ToDecimal(BoltThread.Substring(1));      // It is in mm
                if (UnitSystem_Button.Text == "SAE")                               // Rest of the units such as clamp length are in inches
                {
                    if (Clamp1_TextBox.Text.Length > 0)
                    {
                        TopClamp = Math.Round((Convert.ToDecimal(Clamp1_TextBox.Text) * 25.4M), 3, MidpointRounding.AwayFromZero);      // Convert to mm from inch
                    }

                    if (Clamp2_TextBox.Text.Length > 0)
                    {
                        BottomClamp = Math.Round((Convert.ToDecimal(Clamp2_TextBox.Text) * 25.4M), 3, MidpointRounding.AwayFromZero);   // Convert to mm from inch
                    }
                    if (GasketGap_TextBox.Text.Length > 0)
                    {
                        GasketThickness = Math.Round((Convert.ToDecimal(GasketGap_TextBox.Text) * 25.4M), 3, MidpointRounding.AwayFromZero);    // Convert to mm from inch
                    }
                    if (CustomSpacer_TextBox.Text.Length > 0)
                    {
                        SpacerThickness = Math.Round((Convert.ToDecimal(CustomSpacer_TextBox.Text) * 25.4M), 3, MidpointRounding.AwayFromZero);    // Convert to mm from inch
                    }
                }
                else        // Rest of the units such as clamp length are in mm
                {
                    if (Clamp1_TextBox.Text.Length > 0)
                    {
                        TopClamp = Convert.ToDecimal(Clamp1_TextBox.Text);
                    }
                    if (Clamp2_TextBox.Text.Length > 0)
                    {
                        BottomClamp = Convert.ToDecimal(Clamp2_TextBox.Text);
                    }
                    if (GasketGap_TextBox.Text.Length > 0)
                    {
                        GasketThickness = Convert.ToDecimal(GasketGap_TextBox.Text);
                    }
                    if (CustomSpacer_TextBox.Text.Length > 0)
                    {
                        SpacerThickness = Convert.ToDecimal(CustomSpacer_TextBox.Text);
                    }
                }

            }
            else               // Nominal diameter is in inch
            {
                BoltThread = BoltThread.Substring(0, BoltThread.Length - 1);
                string[] SeparateDigits = BoltThread.Split(' ');
                string[] SeparateFraction;
                decimal CompletePart = 0M;
                decimal Numerator = 0M;
                decimal Denominator = 0M;
                if (SeparateDigits.Count() == 1)
                {
                    SeparateFraction = SeparateDigits[0].Split('/');
                    if (SeparateFraction.Count() > 1)
                    {
                        Numerator = Convert.ToInt32(SeparateFraction[0]);
                        Denominator = Convert.ToInt32(SeparateFraction[1]);
                    }
                }
                else if (SeparateDigits.Count() > 1)
                {
                    SeparateFraction = SeparateDigits[1].Split('/');
                    CompletePart = Convert.ToInt32(SeparateDigits[0]);
                    if (SeparateFraction.Count() > 1)
                    {
                        Numerator = Convert.ToInt32(SeparateFraction[0]);
                        Denominator = Convert.ToInt32(SeparateFraction[1]);
                    }
                    else
                    {
                        MessageBox.Show("Bolt nominal diameter or bolt thread value is wrong.Cannot calculate the LTF.");
                        this.Close();
                    }
                }

                if (Denominator > 0)
                {
                    decimal Ratio = Numerator / Denominator;
                    BoltDiameter = (CompletePart + Math.Round(Ratio, 4, MidpointRounding.AwayFromZero)) * 25.4M;
                }
                else
                {
                    BoltDiameter = (CompletePart + Numerator) * 25.4M;
                }

                if (UnitSystem_Button.Text == "SAE")                               // Rest of the units such as clamp length are in inches
                {
                    if (Clamp1_TextBox.Text.Length > 0)
                    {
                        TopClamp = Math.Round((Convert.ToDecimal(Clamp1_TextBox.Text) * 25.4M), 3, MidpointRounding.AwayFromZero);      // Convert to mm from inch
                    }

                    if (Clamp2_TextBox.Text.Length > 0)
                    {
                        BottomClamp = Math.Round((Convert.ToDecimal(Clamp2_TextBox.Text) * 25.4M), 3, MidpointRounding.AwayFromZero);   // Convert to mm from inch
                    }
                    if (GasketGap_TextBox.Text.Length > 0)
                    {
                        GasketThickness = Math.Round((Convert.ToDecimal(GasketGap_TextBox.Text) * 25.4M), 3, MidpointRounding.AwayFromZero);    // Convert to mm from inch
                    }

                    if (CustomSpacer_TextBox.Text.Length > 0)
                    {
                        SpacerThickness = Math.Round((Convert.ToDecimal(CustomSpacer_TextBox.Text) * 25.4M), 3, MidpointRounding.AwayFromZero);    // Convert to mm from inch
                    }
                }
                else        // Rest of the units such as clamp length are in mm
                {
                    if (Clamp1_TextBox.Text.Length > 0)
                    {
                        TopClamp = Convert.ToDecimal(Clamp1_TextBox.Text);
                    }
                    if (Clamp2_TextBox.Text.Length > 0)
                    {
                        BottomClamp = Convert.ToDecimal(Clamp2_TextBox.Text);
                    }
                    if (GasketGap_TextBox.Text.Length > 0)
                    {
                        GasketThickness = Convert.ToDecimal(GasketGap_TextBox.Text);
                    }

                    if (CustomSpacer_TextBox.Text.Length > 0)
                    {
                        SpacerThickness = Convert.ToDecimal(CustomSpacer_TextBox.Text);
                    }
                }
            }
            // All units are in mm
            CurrentClampLength = TopClamp + BottomClamp + GasketThickness + SpacerThickness;
            if (CurrentClampLength > 0)
            {
                CurrentLTF = 1.01M + Math.Round((BoltDiameter / CurrentClampLength), 4, MidpointRounding.AwayFromZero);
            }
            LTF_MainCurrent = CurrentLTF;
            return CurrentLTF;
        }

        private void BoltMaterial_ComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            decimal Material_BoltDiameter = 0M;
            string Material_Id = BoltMaterial_ComboBox.SelectedValue.ToString();
            if (Material_Id != "System.Data.DataRowView")
            {
                if (BoltThread_ComboBox.SelectedItem != null)
                {
                    string Bolt_Diameter = ((DataRowView)BoltThread_ComboBox.SelectedItem)["BoltThread"].ToString();
                    
                    if (Bolt_Diameter.Substring(0, 1) == "M")
                    {
                        Material_BoltDiameter = Convert.ToDecimal(Bolt_Diameter.Substring(1));
                    }
                    else
                    {
                        Bolt_Diameter = Bolt_Diameter.Substring(0, (Bolt_Diameter.Length - 1)).Trim();
                        Material_BoltDiameter = Convert_Decimal(Bolt_Diameter);
                        Material_BoltDiameter = Math.Round((Material_BoltDiameter * 25.4M), 4, MidpointRounding.AwayFromZero);
                    }
                    if (IsManual)
                    {
                        BoltMaterial_ComboBox.BackColor = Color.White;
                    }
                    BoltMaterial_Info_Label.Text = ((DataRowView)BoltMaterial_ComboBox.SelectedItem)["Material"].ToString();
                    BoltMaterialValueTab_Label.Text = BoltMaterial_Info_Label.Text;

                    // Get Yield from database
                    SQLiteConnection BoltMaterialConnect = new SQLiteConnection("Data Source=" + BasePath + "\\WizBolt.db;New=False;");
                    SQLiteCommand PopulateBoltMaterial_SQLiteCommand = new SQLiteCommand(BoltMaterialConnect);
                    BoltMaterialConnect.Open();

                    PopulateBoltMaterial_SQLiteCommand.CommandText = "SELECT   M.MaterialId,    (S.StandardAbbreviation || ' ' || SA.StandardAuxiliary || '-' || COALESCE(M.MaterialAbbreviation, ' ')) AS  Material, BM.Yield	"
                                                                            + "   FROM Materials 		M "
                                                                            + "   		INNER JOIN StandardAuxiliaries  SA      ON          SA.StandardAuxiliaryId		=		M.StandardAuxiliaryId   "
                                                                            + "   													AND     SA.StandardId 				=		M.StandardId            "
                                                                            + "   		INNER JOIN Standards   			S		ON			S.StandardId				=		M.StandardId            "
                                                                            + "   		INNER JOIN 	BoltMaterial 		BM		ON			BM.MaterialId				=		M.MaterialId            "
                                                                            + "   			WHERE 		SA.StandardId = 1	"
                                                                            + "   						AND SA.StandardAuxiliaryId IN (1,  4)   "
                                                                            + "   						AND M.MaterialId = " + Material_Id 
                                                                            + "   						AND " + Material_BoltDiameter.ToString() + " BETWEEN   BM.Min_BoltDiameter AND BM.Max_BoltDiameter;    ";

                    DataSet BoltMaterial_DataSet = new DataSet();
                    SQLiteDataAdapter BoltMaterial_DataAdapter = new SQLiteDataAdapter(PopulateBoltMaterial_SQLiteCommand);
                    int BoltMaterialResult = BoltMaterial_DataAdapter.Fill(BoltMaterial_DataSet);
                    if (BoltMaterialResult > 0)
                    {
                        BoltYield_SI_Label.Text = BoltMaterial_DataSet.Tables[0].Rows[0]["Yield"].ToString().Trim() + " MPa";
                        BoltYield_In_Label.Text = Math.Round((Convert.ToDecimal(BoltMaterial_DataSet.Tables[0].Rows[0]["Yield"].ToString()) * 0.145037738007M), 4, MidpointRounding.AwayFromZero).ToString().Trim() + " ksi";
                        BoltYieldStrengthValueTab_In_Label.Text = BoltYield_In_Label.Text;
                        BoltYieldStrengthValueTab_SI_Label.Text = BoltYield_SI_Label.Text;
                        DrawGraph();
                    }
                    
                }
            }
        }

        // Public functions for transactions

        public void SaveApplication()
        {
            // Variables for local use
            bool IsUnique = true;
            int MaxAppID = 0;

            // Data values from drop down list
            string SpecificationValue = string.Empty;
            string Flange_Rating = string.Empty;
            string Flange1_AbbreviationValue = string.Empty;
            string Flange2_AbbreviationValue = string.Empty;
            string Bolt_Thread = string.Empty;
            string Model_Number = string.Empty;
            string Material_Value = string.Empty;
            string Pitch_TPI_Title = string.Empty;
            string BoltStressBase = string.Empty;
            string Gasket_Id = string.Empty;
            string Gasket_Name = string.Empty;
            string Flange2_Id = string.Empty;
            string Spacer_Id = string.Empty;
            string Spacer_Name = string.Empty;
            string BoltTool_RatioId = string.Empty;
            string BoltTool_Ratio = string.Empty;
            string Tool_Id = string.Empty;
            string BoltDiameter_Id = string.Empty;
            string Material_Id = string.Empty;
            string FlangeTop_Thickness = string.Empty;
            string FlangeBottom_Thickness = string.Empty;
            string Gasket_Gap = string.Empty;
            string Spacer_Thickness = string.Empty;
            string FirstPass_100 = string.Empty;
            string FirstPass_50 = string.Empty;
            string SecondPass_50 = string.Empty;
            string Max_Detensioning = string.Empty;
            string IsDetensioned = string.Empty;
            string UnitSystem = string.Empty;
            string TorqueValue = string.Empty;
            string ClampLength = string.Empty;
            string PitchValue = string.Empty;
            string BoltNumbers = string.Empty;
            string LTFValue = string.Empty;
            string ToolPressureArea = string.Empty;
            string BoltYieldValue = string.Empty;
            string TensileStressAreaValue = string.Empty;
            string BoltLengthValue = string.Empty;
            string T1ABoltStressValue = string.Empty;
            string T1ABoltLoadValue = string.Empty;
            string T1ABoltYieldValue = string.Empty;
            string T1BBoltStressValue = string.Empty;
            string T1BBoltLoadValue = string.Empty;
            string T1BBoltYieldValue = string.Empty;
            string T2RBoltStressValue = string.Empty;
            string T2RBoltLoadValue = string.Empty;
            string T2RBoltYieldValue = string.Empty;
            string T3RBoltStressValue = string.Empty;
            string T3RBoltLoadValue = string.Empty;
            string T3RBoltYieldValue = string.Empty;
            string DetenBoltStressValue = string.Empty;
            string DetenBoltLoadValue = string.Empty;
            string DetenBoltYieldValue = string.Empty;
            string Pass1AppliedPressure = string.Empty;
            string Pass2AppliedPressure = string.Empty;
            string Pass3AppliedPressure = string.Empty;
            string CheckPass1AppliedPressure = string.Empty;
            string FrictionalCoefficientValue = string.Empty;
            string ResidualStressValue = string.Empty;
            string CrossLoadingValue = string.Empty;
            string DetensioningPerCent = string.Empty;
            //string BoltingImage = string.Empty;

            if ((ResidualBoltStress_TextBox.Text != string.Empty) && (ResidualBoltStress_TextBox.Text.Length > 0))
            {
                if (Convert.ToDecimal(ResidualBoltStress_TextBox.Text) > 0)
                {

                    // Assign values of drop down list as this create problem
                    //((DataRowView)Specification_ComboBox.SelectedItem)["Standard"].ToString();
                    SpecificationValue = ((DataRowView)Specification_ComboBox.SelectedItem)["Standard"].ToString();
                    Flange_Rating = ((DataRowView)Rating_ComboBox.SelectedItem)["AbbreviatedRating"].ToString();
                    Flange1_AbbreviationValue = ((DataRowView)Flange1Config_ComboBox.SelectedItem)["FlangeAbbreviation"].ToString();

                    if (DeTension_CheckBox.Checked)
                    {
                        IsDetensioned = "Yes";
                    }
                    else
                    {
                        IsDetensioned = "No";
                    }

                    if (UnitSystem_Button.Text == "SAE")
                    {
                        UnitSystem = "SAE";

                        if (Clamp1_TextBox.Text.Length > 0)
                        {
                            FlangeTop_Thickness = (Convert.ToDecimal(Clamp1_TextBox.Text) * 25.4M).ToString();
                        }
                        else
                        {
                            FlangeTop_Thickness = "0";
                        }
                        if (GasketGap_TextBox.Text.Length > 0)
                        {
                            Gasket_Gap = (Convert.ToDecimal(GasketGap_TextBox.Text) * 25.4M).ToString();
                        }
                        else
                        {
                            Gasket_Gap = "0";
                        }
                        if (CustomSpacer_TextBox.Text.Length > 0)
                        {
                            Spacer_Thickness = (Convert.ToDecimal(CustomSpacer_TextBox.Text) * 25.4M).ToString();
                        }
                        else
                        {
                            Spacer_Thickness = "0";
                        }
                    }
                    else
                    {
                        UnitSystem = "S. I. Unit";
                        if (Clamp1_TextBox.Text.Length > 0)
                        {
                            FlangeTop_Thickness = Clamp1_TextBox.Text;
                        }
                        else
                        {
                            FlangeTop_Thickness = "0";
                        }
                        if (GasketGap_TextBox.Text.Length > 0)
                        {
                            Gasket_Gap = GasketGap_TextBox.Text;
                        }
                        else
                        {
                            Gasket_Gap = "0";
                        }
                        if (CustomSpacer_TextBox.Text.Length > 0)
                        {
                            Spacer_Thickness = CustomSpacer_TextBox.Text;
                        }
                        else
                        {
                            Spacer_Thickness = "0";
                        }
                    }

                    if (Flange2Config_ComboBox.SelectedItem != null)
                    {
                        Flange2_Id = Flange2Config_ComboBox.SelectedValue.ToString();
                        Flange2_AbbreviationValue = ((DataRowView)Flange2Config_ComboBox.SelectedItem)["FlangeAbbreviation"].ToString();
                        if (UnitSystem_Button.Text == "SAE")
                        {
                            FlangeBottom_Thickness = (Convert.ToDecimal(Clamp2_TextBox.Text) * 25.4M).ToString();
                        }
                        else
                        {
                            FlangeBottom_Thickness = Clamp2_TextBox.Text;
                        }
                    }
                    else
                    {
                        Flange2Config_ComboBox.SelectedValue = "2";
                        Flange2Config_ComboBox.SelectedItem = "WN-RF";
                        Flange2Config_ComboBox.SelectedIndex = 0;
                        Flange2_Id = "2";
                        Flange2_AbbreviationValue = "WN-RF";
                        FlangeBottom_Thickness = "25.4";
                    }
                    if (BoltThread_ComboBox.SelectedItem != null)
                    {
                        BoltDiameter_Id = BoltThread_ComboBox.SelectedValue.ToString();
                        Bolt_Thread = ((DataRowView)BoltThread_ComboBox.SelectedItem)["BoltThread"].ToString();
                    }
                    else
                    {
                        Bolt_Thread = "3/4\"";
                        BoltDiameter_Id = "1";
                    }
                    if (TensionerTool_ComboBox.SelectedItem != null)
                    {
                        Model_Number = ((DataRowView)TensionerTool_ComboBox.SelectedItem)["ModelNumber"].ToString();
                        Tool_Id = TensionerTool_ComboBox.SelectedValue.ToString();
                    }
                    else
                    {
                        Tool_Id = "";
                        Model_Number = "";
                    }
                    if (BoltMaterial_ComboBox.SelectedItem != null)
                    {
                        Material_Id = BoltMaterial_ComboBox.SelectedValue.ToString();
                        Material_Value = ((DataRowView)BoltMaterial_ComboBox.SelectedItem)["Material"].ToString();
                    }
                    else
                    {
                        Material_Value = "ASTM A193 - B7";
                        Material_Id = "3";
                    }
                    if (Gasket_ComboBox.SelectedItem.ToString() != null)
                    {
                        Gasket_Id = (Gasket_ComboBox.SelectedIndex + 1).ToString();
                        Gasket_Name = Gasket_ComboBox.SelectedItem.ToString();
                    }
                    else
                    {
                        Gasket_ComboBox.SelectedItem = "Seal Ring";
                        Gasket_ComboBox.SelectedValue = "1";
                        Gasket_Name = "Seal Ring";
                        Gasket_Id = "1";
                    }

                    if (BoltTool_ComboBox.SelectedItem != null)
                    {
                        BoltTool_RatioId = (BoltTool_ComboBox.SelectedIndex + 1).ToString();
                        BoltTool_Ratio = BoltTool_ComboBox.SelectedItem.ToString();
                    }
                    else
                    {
                        BoltTool_Ratio = "100%";
                        BoltTool_ComboBox.SelectedItem = "100%";
                        BoltTool_ComboBox.SelectedValue = "1";
                        BoltTool_RatioId = "1";
                    }
                    if (Bolt_Thread.Substring(0, 1) == "M")
                    {
                        Pitch_TPI_Title = "Pitch";
                    }
                    else
                    {
                        Pitch_TPI_Title = "TPI";
                    }
                    if (TensileStressArea_RadioButton.Checked)
                    {
                        BoltStressBase = "Tensile Stress Area";
                    }
                    else
                    {
                        BoltStressBase = "Minor Diameter Area";
                    }

                    // Treating null values. Only numberic nulls need treatment as string values has single quotes i.e. ''
                    if (Spacer_ComboBox.SelectedItem != null)
                    {
                        Spacer_Id = (Spacer_ComboBox.SelectedIndex + 1).ToString();
                        Spacer_Name = Spacer_ComboBox.SelectedItem.ToString();
                    }
                    else
                    {
                        Spacer_Id = "NULL";
                        Spacer_Name = "";
                    }
                    if ((CustomSpacer_TextBox.Text.Length == 0) && (CustomSpacer_TextBox.Text == ""))
                    {
                        CustomSpacer_TextBox.Text = "NULL";
                    }
                    
                     int Bolt_ToolId = BoltTool_ComboBox.SelectedIndex;
                     switch (Bolt_ToolId)
                     {
                         case 0:
                             if ((PressureB_SI_Label.Text.Trim() != "") && (PressureB_SI_Label.Text.Trim().Length > 4))
                             {
                                 FirstPass_100 = PressureB_SI_Label.Text.Substring(0, PressureB_SI_Label.Text.Length - 4);
                                 Max_Detensioning = FirstPass_100;
                             }
                             else
                             {
                                 FirstPass_100 = "NULL";
                                 Max_Detensioning = "NULL";
                             }
                             
                                 FirstPass_50 = "NULL";
                                 SecondPass_50 = "NULL";

                             break;
                         case 1:
                             FirstPass_100 = "NULL";
                             if ((PressureA_SI_Label.Text.Trim() != "") && (PressureA_SI_Label.Text.Trim().Length > 4))
                             {
                                 FirstPass_50 = PressureA_SI_Label.Text.Substring(0, PressureA_SI_Label.Text.Length - 4);
                             }
                             else
                             {
                                 FirstPass_50 = "NULL";
                             }
                             if ((PressureB_SI_Label.Text.Trim() != "") && (PressureB_SI_Label.Text.Trim().Length > 4))
                             {
                                 SecondPass_50 = PressureB_SI_Label.Text.Substring(0, PressureB_SI_Label.Text.Length - 4);
                                 Max_Detensioning = SecondPass_50;
                             }
                             else
                             {
                                 SecondPass_50 = "NULL";
                                 Max_Detensioning = SecondPass_50;
                             }

                             break;
                         case 2:
                             FirstPass_100 = "NULL";
                             if ((PressureA_SI_Label.Text.Trim() != "") && (PressureA_SI_Label.Text.Trim().Length > 4))
                             {
                                 FirstPass_50 = PressureA_SI_Label.Text.Substring(0, PressureA_SI_Label.Text.Length - 4);
                             }
                             else
                             {
                                 FirstPass_50 = "NULL";
                             }
                             if ((PressureB_SI_Label.Text.Trim() != "") && (PressureB_SI_Label.Text.Trim().Length > 4))
                             {
                                 SecondPass_50 = PressureB_SI_Label.Text.Substring(0, PressureB_SI_Label.Text.Length - 4);
                                 Max_Detensioning = SecondPass_50;
                             }
                             else
                             {
                                 SecondPass_50 = "NULL";
                                 Max_Detensioning = SecondPass_50;
                             }

                             break;
                         case 3:
                             FirstPass_100 = "NULL";
                             FirstPass_50 = "NULL";
                             SecondPass_50 = "NULL";
                             Max_Detensioning = SecondPass_50;
                             break;
                         default:
                             FirstPass_100 = "NULL";
                             FirstPass_50 = "NULL";
                             SecondPass_50 = "NULL";
                             Max_Detensioning = SecondPass_50;
                             break;
                     }
                    TorqueValue = Torque_SI_Label.Text.Substring(0, Torque_SI_Label.Text.Length - 4);
                    if ((TorqueValue == "") && (TorqueValue == string.Empty) && (TorqueValue.Length == 0))
                    {
                        TorqueValue = "NULL";
                    }
                    else
                    {
                        try
                        {
                            decimal TorqueTest = Convert.ToDecimal(TorqueValue);
                        }
                        catch (Exception Torque_Exception)
                        {
                            MessageBox.Show(Torque_Exception.Message + " torque value is wrong. Only numerals allowed.");
                        }
                    }
                    ClampLength = ClampLengthInfo_mm_Label.Text.Substring(0, ClampLengthInfo_mm_Label.Text.Length - 3);
                    if ((ClampLength == "") && (ClampLength == string.Empty) && (ClampLength.Length == 0))
                    {
                        ClampLength = "NULL";
                    }
                    else
                    {
                        try
                        {
                            decimal ClampTest = Convert.ToDecimal(ClampLength);
                        }
                        catch (Exception Clamp_Exception)
                        {
                            MessageBox.Show(Clamp_Exception.Message + " total clamp length is wrong. Only numerals allowed.");
                        }
                    }

                    PitchValue = Pitch_TextBox.Text;
                    if ((PitchValue == "") && (PitchValue == string.Empty) && (PitchValue.Length == 0))
                    {
                        PitchValue = "NULL";
                    }
                    else
                    {
                        try
                        {
                            decimal PitchTest = Convert.ToDecimal(PitchValue);
                        }
                        catch (Exception Pitch_Exception)
                        {
                            MessageBox.Show(Pitch_Exception.Message + " Pitch/TPI is wrong. Only numerals allowed.");
                        }
                    }

                    BoltNumbers = NumberOfBolts_TextBox.Text;
                    if ((BoltNumbers == "") && (BoltNumbers == string.Empty) && (BoltNumbers.Length == 0))
                    {
                        BoltNumbers = "NULL";
                    }
                    else
                    {
                        try
                        {
                            decimal BoltNumberTest = Convert.ToDecimal(BoltNumbers);
                        }
                        catch (Exception BoltNumber_Exception)
                        {
                            MessageBox.Show(BoltNumber_Exception.Message + " number of bolts value is wrong. Only numerals allowed.");
                        }
                    }

                    LTFValue = LTF_Label.Text;
                    if ((LTFValue == "") && (LTFValue == string.Empty) && (LTFValue.Length == 0))
                    {
                        LTFValue = "NULL";
                    }
                    else
                    {
                        try
                        {
                            decimal LTFValueTest = Convert.ToDecimal(LTFValue);
                        }
                        catch (Exception LTFValue_Exception)
                        {
                            MessageBox.Show(LTFValue_Exception.Message + " Load Transfer Factor is wrong. Only numerals allowed.");
                        }
                    }
                    if ((ToolPressureArea_mm_Label.Text.Trim() != "") && (ToolPressureArea_mm_Label.Text.Trim().Length > 5))
                    {
                        ToolPressureArea = ToolPressureArea_mm_Label.Text.Substring(0, ToolPressureArea_mm_Label.Text.Length - 5);
                    }
                    if ((ToolPressureArea == "") && (ToolPressureArea == string.Empty) && (ToolPressureArea.Length == 0))
                    {
                        ToolPressureArea = "NULL";
                    }
                    else
                    {
                        try
                        {
                            decimal ToolPressureAreaTest = Convert.ToDecimal(ToolPressureArea);
                        }
                        catch (Exception ToolPressureArea_Exception)
                        {
                            MessageBox.Show(ToolPressureArea_Exception.Message + " Tool Pressure Area is wrong. Only numerals allowed.");
                        }
                    }

                    BoltYieldValue = BoltYield_SI_Label.Text.Substring(0, BoltYield_SI_Label.Text.Length - 4);
                    if ((BoltYieldValue == "") && (BoltYieldValue == string.Empty) && (BoltYieldValue.Length == 0))
                    {
                        BoltYieldValue = "NULL";
                    }
                    else
                    {
                        try
                        {
                            decimal BoltYieldValueTest = Convert.ToDecimal(BoltYieldValue);
                        }
                        catch (Exception BoltYield_Exception)
                        {
                            MessageBox.Show(BoltYield_Exception.Message + " Bolt Yield Value is wrong. Only numerals allowed.");
                        }
                    }

                    TensileStressAreaValue = TensileStressArea_SI_Label.Text.Substring(0, TensileStressArea_SI_Label.Text.Length - 5);
                    if ((TensileStressAreaValue == "") && (TensileStressAreaValue == string.Empty) && (TensileStressAreaValue.Length == 0))
                    {
                        TensileStressAreaValue = "NULL";
                    }
                    else
                    {
                        try
                        {
                            decimal TensileStressAreaValueTest = Convert.ToDecimal(TensileStressAreaValue);
                        }
                        catch (Exception TensileStressArea_Exception)
                        {
                            MessageBox.Show(TensileStressArea_Exception.Message + " Tensile Stress Area is wrong. Only numerals allowed.");
                        }
                    }

                    BoltLengthValue = BoltLength_SI_Label.Text.Substring(0, BoltLength_SI_Label.Text.Length - 3);
                    if ((BoltLengthValue == "") && (BoltLengthValue == string.Empty) && (BoltLengthValue.Length == 0))
                    {
                        BoltLengthValue = "NULL";
                    }
                    else
                    {
                        try
                        {
                            decimal BoltLengthTest = Convert.ToDecimal(BoltLengthValue);
                        }
                        catch (Exception BoltLength_Exception)
                        {
                            MessageBox.Show(BoltLength_Exception.Message + " Bolt Length is wrong. Only numerals allowed.");
                        }
                    }

                    T1ABoltStressValue = T1ABoltStress_SI_Label.Text;
                    if ((T1ABoltStressValue == "") && (T1ABoltStressValue == string.Empty) && (T1ABoltStressValue.Length == 0))
                    {
                        T1ABoltStressValue = "NULL";
                    }
                    else
                    {
                        try
                        {
                            decimal T1ABoltStressTest = Convert.ToDecimal(T1ABoltStressValue);
                        }
                        catch (Exception T1ABoltStress_Exception)
                        {
                            MessageBox.Show(T1ABoltStress_Exception.Message + " Higher Bolt Stress is wrong. Only numerals allowed.");
                        }
                    }

                    T1ABoltLoadValue = T1ABoltLoad_SI_Label.Text;
                    if ((T1ABoltLoadValue == "") && (T1ABoltLoadValue == string.Empty) && (T1ABoltLoadValue.Length == 0))
                    {
                        T1ABoltLoadValue = "NULL";
                    }
                    else
                    {
                        try
                        {
                            decimal T1ABoltLoadTest = Convert.ToDecimal(T1ABoltLoadValue);
                        }
                        catch (Exception T1ABoltLoad_Exception)
                        {
                            MessageBox.Show(T1ABoltLoad_Exception.Message + " Higher Bolt Load value is wrong. Only numerals allowed.");
                        }
                    }

                    T1ABoltYieldValue = T1ABoltYield_SI_Label.Text;
                    if ((T1ABoltYieldValue == "") && (T1ABoltYieldValue == string.Empty) && (T1ABoltYieldValue.Length == 0))
                    {
                        T1ABoltYieldValue = "NULL";
                    }
                    else
                    {
                        try
                        {
                            decimal T1ABoltYieldTest = Convert.ToDecimal(T1ABoltYieldValue);
                        }
                        catch (Exception T1ABoltYield_Exception)
                        {
                            MessageBox.Show(T1ABoltYield_Exception.Message + " Higher Pressure Bolt Yield per cent value is wrong. Only numerals allowed.");
                        }
                    }

                    T1BBoltStressValue = T1BBoltStress_SI_Label.Text;
                    if ((T1BBoltStressValue == "") && (T1BBoltStressValue == string.Empty) && (T1BBoltStressValue.Length == 0))
                    {
                        T1BBoltStressValue = "NULL";
                    }
                    else
                    {
                        try
                        {
                            decimal T1BBoltStressTest = Convert.ToDecimal(T1BBoltStressValue);
                        }
                        catch (Exception T1BBoltStress_Exception)
                        {
                            MessageBox.Show(T1BBoltStress_Exception.Message + " Bolt Stress is wrong. Only numerals allowed.");
                        }
                    }

                    T1BBoltLoadValue = T1ABoltLoad_SI_Label.Text;
                    if ((T1BBoltLoadValue == "") && (T1BBoltLoadValue == string.Empty) && (T1BBoltLoadValue.Length == 0))
                    {
                        T1BBoltLoadValue = "NULL";
                    }
                    else
                    {
                        try
                        {
                            decimal T1BBoltLoadTest = Convert.ToDecimal(T1BBoltLoadValue);
                        }
                        catch (Exception T1BBoltLoad_Exception)
                        {
                            MessageBox.Show(T1BBoltLoad_Exception.Message + " Bolt Load is wrong. Only numerals allowed.");
                        }
                    }

                    T1BBoltYieldValue = T1BBoltYield_Per_Label.Text;
                    if ((T1BBoltYieldValue == "") && (T1BBoltYieldValue == string.Empty) && (T1BBoltYieldValue.Length == 0))
                    {
                        T1BBoltYieldValue = "NULL";
                    }
                    else
                    {
                        try
                        {
                            decimal T1BBoltYieldTest = Convert.ToDecimal(T1BBoltYieldValue);
                        }
                        catch (Exception T1BBoltYield_Exception)
                        {
                            MessageBox.Show(T1BBoltYield_Exception.Message + " Bolt Yield per cent is wrong. Only numerals allowed.");
                        }
                    }

                    T2RBoltStressValue = T2RBoltStress_SI_Label.Text;
                    if ((T2RBoltStressValue == "") && (T2RBoltStressValue == string.Empty) && (T2RBoltStressValue.Length == 0))
                    {
                        T2RBoltStressValue = "NULL";
                    }
                    else
                    {
                        try
                        {
                            decimal T2RBoltStressTest = Convert.ToDecimal(T2RBoltStressValue);
                        }
                        catch (Exception T2RBoltStress_Exception)
                        {
                            MessageBox.Show(T2RBoltStress_Exception.Message + " Bolt stress (normal) is wrong. Only numerals allowed.");
                        }
                    }

                    T2RBoltLoadValue = T2RBoltLoad_SI_Label.Text;
                    if ((T2RBoltLoadValue == "") && (T2RBoltLoadValue == string.Empty) && (T2RBoltLoadValue.Length == 0))
                    {
                        T2RBoltLoadValue = "NULL";
                    }
                    else
                    {
                        try
                        {
                            decimal T2RBoltLoadTest = Convert.ToDecimal(T2RBoltLoadValue);
                        }
                        catch (Exception T2RBoltLoad_Exception)
                        {
                            MessageBox.Show(T2RBoltLoad_Exception.Message + " Bolt load (normal) is wrong. Only numerals allowed.");
                        }
                    }

                    T2RBoltYieldValue = T2RBoltYield_Per_Label.Text;
                    if ((T2RBoltYieldValue == "") && (T2RBoltYieldValue == string.Empty) && (T2RBoltYieldValue.Length == 0))
                    {
                        T2RBoltYieldValue = "NULL";
                    }
                    else
                    {
                        try
                        {
                            decimal T2RBoltYield = Convert.ToDecimal(T2RBoltYieldValue);
                        }
                        catch (Exception T2RBoltYield_Exception)
                        {
                            MessageBox.Show(T2RBoltYield_Exception.Message + " Bolt yield per cent (normal) is wrong. Only numerals allowed.");
                        }
                    }
                    
                    T3RBoltStressValue = T3RBoltStress_SI_Label.Text;
                    if ((T3RBoltStressValue == "") && (T3RBoltStressValue == string.Empty) && (T3RBoltStressValue.Length == 0))
                    {
                        T3RBoltStressValue = "NULL";
                    }
                    else
                    {
                        try
                        {
                            decimal T3RBoltStressTest = Convert.ToDecimal(T3RBoltStressValue);
                        }
                        catch (Exception T3RBoltStress_Exception)
                        {
                            MessageBox.Show(T3RBoltStress_Exception.Message + " Bolt stress (normal) is wrong. Only numerals allowed.");
                        }
                    }

                    T3RBoltLoadValue = T3RBoltLoad_SI_Label.Text;
                    if ((T3RBoltLoadValue == "") && (T3RBoltLoadValue == string.Empty) && (T3RBoltLoadValue.Length == 0))
                    {
                        T3RBoltLoadValue = "NULL";
                    }
                    else
                    {
                        try
                        {
                            decimal T3RBoltLoadTest = Convert.ToDecimal(T3RBoltLoadValue);
                        }
                        catch (Exception T3RBoltLoad_Exception)
                        {
                            MessageBox.Show(T3RBoltLoad_Exception.Message + " Bolt load (normal) is wrong. Only numerals allowed.");
                        }
                    }

                    T3RBoltYieldValue = T3RBoltYield_Per_Label.Text;
                    if ((T3RBoltYieldValue == "") && (T3RBoltYieldValue == string.Empty) && (T3RBoltYieldValue.Length == 0))
                    {
                        T3RBoltYieldValue = "NULL";
                    }
                    else
                    {
                        try
                        {
                            decimal T3RBoltYield = Convert.ToDecimal(T3RBoltYieldValue);
                        }
                        catch (Exception T3RBoltYield_Exception)
                        {
                            MessageBox.Show(T3RBoltYield_Exception.Message + " Bolt yield per cent (normal) is wrong. Only numerals allowed.");
                        }
                    }

                    DetenBoltStressValue = DetenBoltStress_SI_Label.Text;
                    if ((DetenBoltStressValue == "") && (DetenBoltStressValue == string.Empty) && (DetenBoltStressValue.Length == 0))
                    {
                        DetenBoltStressValue = "NULL";
                    }
                    else
                    {
                        try
                        {
                            decimal DetenBoltStressTest = Convert.ToDecimal(DetenBoltStressValue);
                        }
                        catch (Exception DetenBoltStress_Exception)
                        {
                            MessageBox.Show(DetenBoltStress_Exception.Message + " De-tensioning bolt stress is wrong. Only numerals allowed.");
                        }
                    }

                    DetenBoltLoadValue = DetenBoltLoad_SI_Label.Text;
                    if ((DetenBoltLoadValue == "") && (DetenBoltLoadValue == string.Empty) && (DetenBoltLoadValue.Length == 0))
                    {
                        DetenBoltLoadValue = "NULL";
                    }
                    else
                    {
                        try
                        {
                            decimal DetenBoltLoadTest = Convert.ToDecimal(DetenBoltLoadValue);
                        }
                        catch (Exception DetenBoltLoad_Exception)
                        {
                            MessageBox.Show(DetenBoltLoad_Exception.Message + " De-tensioning bolt load is wrong. Only numerals allowed.");
                        }
                    }

                    DetenBoltYieldValue = DetenBoltYield_Per_Label.Text;
                    if ((DetenBoltYieldValue == "") && (DetenBoltYieldValue == string.Empty) && (DetenBoltYieldValue.Length == 0))
                    {
                        DetenBoltYieldValue = "NULL";
                    }
                    else
                    {
                        try
                        {
                            decimal DetenBoltYieldTest = Convert.ToDecimal(DetenBoltYieldValue);
                        }
                        catch (Exception DetenBoltYield_Exception)
                        {
                            MessageBox.Show(DetenBoltYield_Exception.Message + " De-tensioning bolt yield per cent is wrong. Only numerals allowed.");
                        }
                    }

                    Pass1AppliedPressure = Pass1AppliedPressure_SequenceTabValue_SI_Label.Text;
                    if ((Pass1AppliedPressure == "") && (Pass1AppliedPressure == string.Empty) && (Pass1AppliedPressure.Length == 0))
                    {
                        Pass1AppliedPressure = "NULL";
                    }
                    else
                    {
                        try
                        {
                            decimal Pass1AppliedPressureTest = Convert.ToDecimal(Pass1AppliedPressure);
                        }
                        catch (Exception Pass1AppliedPressure_Exception)
                        {
                            MessageBox.Show(Pass1AppliedPressure_Exception.Message + " First pass applied perssure is wrong. Only numerals allowed.");
                        }
                    }

                    Pass2AppliedPressure = Pass2AppliedPressure_SequenceTabValue_SI_Label.Text;
                    if ((Pass2AppliedPressure == "") && (Pass2AppliedPressure == string.Empty) && (Pass2AppliedPressure.Length == 0))
                    {
                        Pass2AppliedPressure = "NULL";
                    }
                    else
                    {
                        try
                        {
                            decimal Pass2AppliedPressureTest = Convert.ToDecimal(Pass2AppliedPressure);
                        }
                        catch (Exception Pass2AppliedPressure_Exception)
                        {
                            MessageBox.Show(Pass2AppliedPressure_Exception.Message + " Second pass applied pressure is wrong. Only numerals allowed.");
                        }
                    }

                    Pass3AppliedPressure = Pass3AppliedPressure_SequenceTabValue_SI_Label.Text;
                    if ((Pass3AppliedPressure == "") && (Pass3AppliedPressure == string.Empty) && (Pass3AppliedPressure.Length == 0))
                    {
                        Pass3AppliedPressure = "NULL";
                    }
                    else
                    {
                        try
                        {
                            decimal Pass3AppliedPressureTest = Convert.ToDecimal(Pass3AppliedPressure);
                        }
                        catch (Exception Pass3AppliedPressure_Exception)
                        {
                            MessageBox.Show(Pass3AppliedPressure_Exception.Message + " Third pass applied pressure is wrong. Only numerals allowed.");
                        }
                    }

                    CheckPass1AppliedPressure = CheckPass1AppliedPressure_SequenceTabValue_SI_Label.Text;
                    if ((CheckPass1AppliedPressure == "") && (CheckPass1AppliedPressure == string.Empty) && (CheckPass1AppliedPressure.Length == 0))
                    {
                        CheckPass1AppliedPressure = "NULL";
                    }
                    else
                    {
                        try
                        {
                            decimal CheckPass1AppliedPressureTest = Convert.ToDecimal(CheckPass1AppliedPressure);
                        }
                        catch (Exception CheckPass1AppliedPressure_Exception)
                        {
                            MessageBox.Show(CheckPass1AppliedPressure_Exception.Message + " Check pass applied pressure is wrong. Only numerals allowed.");
                        }
                    }

                    FrictionalCoefficientValue = CoefficientValue_Label.Text;
                    if ((FrictionalCoefficientValue == "") && (FrictionalCoefficientValue == string.Empty) && (FrictionalCoefficientValue.Length == 0))
                    {
                        FrictionalCoefficientValue = "NULL";
                    }
                    else
                    {
                        try
                        {
                            decimal FrictionalCoefficientTest = Convert.ToDecimal(FrictionalCoefficientValue);
                        }
                        catch (Exception FrictionalCoefficient_Exception)
                        {
                            MessageBox.Show(FrictionalCoefficient_Exception.Message + " Coefficient of friction is wrong. Only numerals allowed.");
                        }
                    }

                    CrossLoadingValue = CrossLoading_TextBox.Text;
                    if ((CrossLoadingValue == "") && (CrossLoadingValue == string.Empty) && (CrossLoadingValue.Length == 0))
                    {
                        CrossLoadingValue = "NULL";
                    }
                    else
                    {
                        try
                        {
                            decimal CrossLoadingTest = Convert.ToDecimal(CrossLoadingValue);
                        }
                        catch (Exception CrossLoadingValue_Exception)
                        {
                            MessageBox.Show(CrossLoadingValue_Exception.Message + " Cross loading per cent is wrong. Only numerals allowed.");
                        }
                    }

                    DetensioningPerCent = Detensioning_TextBox.Text;
                    if ((DetensioningPerCent == "") && (DetensioningPerCent == string.Empty) && (DetensioningPerCent.Length == 0))
                    {
                        DetensioningPerCent = "NULL";
                    }
                    else
                    {
                        try
                        {
                            decimal DetensioningPerCentTest = Convert.ToDecimal(DetensioningPerCent);
                        }
                        catch (Exception DetensioningPerCent_Exception)
                        {
                            MessageBox.Show(DetensioningPerCent_Exception.Message + " De-tensioning per cent is wrong. Only numerals allowed.");
                        }
                    }

                    ResidualStressValue = ResidualBoltStress_TextBox.Text.Trim();
                    if ((ResidualStressValue == "") && (ResidualStressValue == string.Empty) && (ResidualStressValue.Length == 0))
                    {
                        ResidualStressValue = "NULL";
                    }
                    else
                    {
                        try
                        {
                            decimal ResidualStressTest = Convert.ToDecimal(ResidualStressValue);
                        }
                        catch (Exception ResidualStress_Exception)
                        {
                            MessageBox.Show(ResidualStress_Exception.Message + " Residual stress value is wrong. Only numerals allowed.");
                        }
                    }

                    SQLiteConnection ConnectApplication = new SQLiteConnection("Data Source=" + BasePath + "\\WizBolt.db;New=False;");
                    SQLiteCommand PopulateApplication_SQLiteCommand = new SQLiteCommand(ConnectApplication);
                    ConnectApplication.Open();
                    try
                    {
                        PopulateApplication_SQLiteCommand.CommandText = "SELECT 	CustomerId, CustomerName, ProjectReportId, ProjectName, ProjectReference "
                                                                        + " FROM ProjectReport "
                                                                        + "     WHERE ProjectReportId > 0;";
                        DataSet Project_DataSet = new DataSet();
                        SQLiteDataAdapter Project_DataAdapter = new SQLiteDataAdapter(PopulateApplication_SQLiteCommand);
                        int ProjectResult = Project_DataAdapter.Fill(Project_DataSet);
                        if (ProjectResult > 0)
                        {
                            PopulateApplication_SQLiteCommand.CommandText = "SELECT 	P.CustomerId, P.CustomerName, P.ProjectReportId, P.ProjectName, P.ProjectReference, "
                                                                            + " PD.ApplicationId, PD.JointId, PD.PrimaryStandardId, PD.Specification "
                                                                            + " 	FROM ProjectReport	P   "
                                                                            + " 		INNER JOIN	ProjectDetailedReport	PD		ON		PD.ProjectReportId	=	P.ProjectReportId "
                                                                            + "             WHERE P.ProjectReportId > 0 AND P.ProjectReportId = " + ProjectId_Label.Text + "; ";

                            DataSet ApplicationVerifier_DataSet = new DataSet();
                            SQLiteDataAdapter ApplicationVerifier_DataAdapter = new SQLiteDataAdapter(PopulateApplication_SQLiteCommand);
                            int VerifierResult = ApplicationVerifier_DataAdapter.Fill(ApplicationVerifier_DataSet);
                            if (VerifierResult > 0)
                            {
                                for (int AppId = 0; AppId < ApplicationVerifier_DataSet.Tables[0].Rows.Count; AppId++)
                                {
                                    if ((CustomerName_Label.Text == ApplicationVerifier_DataSet.Tables[0].Rows[AppId]["CustomerName"].ToString())
                                        && (ProjectName_Label.Text == ApplicationVerifier_DataSet.Tables[0].Rows[AppId]["ProjectName"].ToString())
                                        && (Reference_Label.Text == ApplicationVerifier_DataSet.Tables[0].Rows[AppId]["ProjectReference"].ToString())
                                        && (JointId_TextBox.Text == ApplicationVerifier_DataSet.Tables[0].Rows[AppId]["JointId"].ToString()))
                                    {
                                        IsUnique = false;
                                    }
                                }
                                if (IsUnique)
                                {
                                    PopulateApplication_SQLiteCommand.CommandText = " SELECT MAX(COALESCE(ApplicationId, 0)) FROM ProjectDetailedReport WHERE ProjectReportId = " + ProjectId_Label.Text + "; ";
                                    SQLiteDataReader App_DataReader = PopulateApplication_SQLiteCommand.ExecuteReader();
                                    if (App_DataReader.Read())
                                    {
                                        if (App_DataReader.IsDBNull(0))
                                        {
                                            MaxAppID = 0;
                                        }
                                        else
                                        {
                                            MaxAppID = App_DataReader.GetInt32(0);
                                        }
                                    }
                                    if (!App_DataReader.IsClosed)
                                    {
                                        App_DataReader.Close();
                                        App_DataReader.Dispose();
                                    }
                                    PopulateApplication_SQLiteCommand.CommandText = "INSERT INTO ProjectDetailedReport (ProjectReportId, ApplicationId, UnitSystem, JointId, PrimaryStandardId,  "
                                               + "   Specification, PrimaryFlangeRatingId, FlangeRating, Flange1_TypeId,  "
                                               + "   Flange1_Abbreviation, Flange1_ClampLength, Flange2_TypeId, Flange2_Abbreviation,  "
                                               + "   Flange2_ClampLength, GasketId, Gasket, GasketGap,  "
                                               + "   SpacerId, Spacer, SpacerThickness, TotalClampLength, PrimaryFlangeBoltId,  "
                                               + "   BoltThread, Pitch_or_TPI, Pitch_TPI_Value, Number_of_Bolts, LTF, "
                                               + "   Bolt_to_ToolRatioId, Bolt_to_ToolRatio, ToolId, ModelNumber,  "
                                               + "   ToolPressureArea, MaterialId, Material, BoltYield, BoltTensileStressArea, "
                                               + "   BoltLength, BoltStressBase, BoltMinorDiameterArea, T1_APressureBoltStress, T1_APressureBoltLoad,  "
                                               + "   T1_APressureBoltYieldPC, T1_BPressureBoltStress, T1_BPressureBoltLoad, T1_BPressureBoltYieldPC,   "
                                               + "   T2_ResidualBoltStress, T2_ResidualBoltLoad, T2_ResidualBoltYieldPC, T3_ResidualBoltStress, T3_ResidualBoltLoad, T3_ResidualBoltYieldPC, "
                                               + "   DetensioningBoltStress, DetensioningBoltLoad,	"
                                               + "   DetensioningBoltYieldPC, TensionPressure_FirstPass, TensionPressure_SecondPass, TensionPressure_ThirdPass, CheckingPass,   "
                                               + "   Torque, Coefficient_Friction, Bolt01, Bolt02, Bolt03, Bolt04, "
                                               + "   FirstPass_100, FirstPass_50, SecondPass_50, Max_Detensioning, "
                                               + "   ResidualBoltStress, Comments, IsDetension, CrossLoadingPC, DetensioningPC"
                                               + ") "
                                               + " VALUES (" + ProjectId_Label.Text + ", " + (MaxAppID + 1).ToString() + ", '" + UnitSystem + "', '" + JointId_TextBox.Text + "', '" + Specification_ComboBox.SelectedValue.ToString() + "', "
                                               + "'" + SpecificationValue + "', '" + Rating_ComboBox.SelectedValue.ToString() + "', '" + Flange_Rating + "', " + Flange1Config_ComboBox.SelectedValue.ToString() + ", "
                                               + "'" + Flange1_AbbreviationValue + "', " + FlangeTop_Thickness + ", " + Flange2_Id + ", '" + Flange2_AbbreviationValue + "', "
                                               + FlangeBottom_Thickness + ", " + Gasket_Id + ", '" + Gasket_Name + "', " + Gasket_Gap + ", "
                                               + Spacer_Id + ", '" + Spacer_Name + "', " + Spacer_Thickness + ",  " + ClampLength + ", '" + BoltDiameter_Id + "', "
                                               + "'" + Bolt_Thread + "', '" + Pitch_TPI_Title + "', " + PitchValue + ", " + BoltNumbers + ", " + LTFValue + ", "
                                               + BoltTool_RatioId + ", '" + BoltTool_Ratio + "', '" + Tool_Id + "', '" + Model_Number + "', "
                                               + ToolPressureArea + ", " + Material_Id + ", '" + Material_Value + "', " + BoltYieldValue + ", " + TensileStressAreaValue + ", "
                                               + BoltLengthValue + ", '" + BoltStressBase + "', " + GlobalMinorDiameterArea + ", " + T1ABoltStressValue + ", " + T1ABoltLoadValue + ", "
                                               + T1ABoltYieldValue + ", " + T1BBoltStressValue + ", " + T1BBoltLoadValue + ", " + T1BBoltYieldValue + ", "
                                               + T2RBoltStressValue + ", " + T2RBoltLoadValue + ", " + T2RBoltYieldValue + ", " + T3RBoltStressValue + ", " + T3RBoltLoadValue + ", " + T3RBoltYieldValue + ", " 
                                               + DetenBoltStressValue + ", " + DetenBoltLoadValue + ", "
                                               + DetenBoltYieldValue + ", " + Pass1AppliedPressure + ", " + Pass2AppliedPressure + ", " + Pass3AppliedPressure + ", " + CheckPass1AppliedPressure + ", "
                                               + TorqueValue + ", " + FrictionalCoefficientValue + ", " + "'1', '" + Pass2Bolt_SequenceTabValue_Label.Text + "', '" + Pass3Bolt_SequenceTabValue_Label.Text + "', '" + Pass4Bolt_SequenceTabValue_Label.Text + "',"
                                               + FirstPass_100 + ", " + FirstPass_50 + ", " + SecondPass_50 + ", " + Max_Detensioning + ", "
                                               + ResidualStressValue + ", '" + Comment_TextBox.Text + "', '" + IsDetensioned + "', " + CrossLoadingValue + ", " + DetensioningPerCent
                                               + ");";

                                    int UniqueAppResult = PopulateApplication_SQLiteCommand.ExecuteNonQuery();
                                }
                                else
                                {
                                    MessageBox.Show("Joint ID is repeated. Every joint must have unique id. This application cannot be saved. Change Joint ID and save again.", "Unique id violated!");
                                }
                            }
                            else
                                if (VerifierResult == 0)
                                {
                                    PopulateApplication_SQLiteCommand.CommandText = "INSERT INTO ProjectDetailedReport (ProjectReportId, ApplicationId, UnitSystem, JointId, PrimaryStandardId,  "
                                               + "   Specification, PrimaryFlangeRatingId, FlangeRating, Flange1_TypeId,  "
                                               + "   Flange1_Abbreviation, Flange1_ClampLength, Flange2_TypeId, Flange2_Abbreviation,  "
                                               + "   Flange2_ClampLength, GasketId, Gasket, GasketGap,  "
                                               + "   SpacerId, Spacer, SpacerThickness, TotalClampLength, PrimaryFlangeBoltId,  "
                                               + "   BoltThread, Pitch_or_TPI, Pitch_TPI_Value, Number_of_Bolts, LTF, "
                                               + "   Bolt_to_ToolRatioId, Bolt_to_ToolRatio, ToolId, ModelNumber,  "
                                               + "   ToolPressureArea, MaterialId, Material, BoltYield, BoltTensileStressArea, "
                                               + "   BoltLength, BoltStressBase, BoltMinorDiameterArea, T1_APressureBoltStress, T1_APressureBoltLoad,  "
                                               + "   T1_APressureBoltYieldPC, T1_BPressureBoltStress, T1_BPressureBoltLoad, T1_BPressureBoltYieldPC,   "
                                               + "   T2_ResidualBoltStress, T2_ResidualBoltLoad, T2_ResidualBoltYieldPC, T3_ResidualBoltStress, T3_ResidualBoltLoad, T3_ResidualBoltYieldPC, "
                                               + "   DetensioningBoltStress, DetensioningBoltLoad,	"
                                               + "   DetensioningBoltYieldPC, TensionPressure_FirstPass, TensionPressure_SecondPass, TensionPressure_ThirdPass, CheckingPass,   "
                                               + "   Torque, Coefficient_Friction, Bolt01, Bolt02, Bolt03, Bolt04,"
                                               + "   FirstPass_100, FirstPass_50, SecondPass_50, Max_Detensioning, "
                                               + "   ResidualBoltStress, Comments, IsDetension, CrossLoadingPC, DetensioningPC"
                                               + ") "
                                               + " VALUES (" + ProjectId_Label.Text + ", " + (MaxAppID + 1).ToString() + ", '" + UnitSystem + "', '" + JointId_TextBox.Text + "', '" + Specification_ComboBox.SelectedValue.ToString() + "', "
                                               + "'" + SpecificationValue + "', '" + Rating_ComboBox.SelectedValue.ToString() + "', '" + Flange_Rating + "', " + Flange1Config_ComboBox.SelectedValue.ToString() + ", "
                                               + "'" + Flange1_AbbreviationValue + "', " + FlangeTop_Thickness + ", " + Flange2_Id + ", '" + Flange2_AbbreviationValue + "', "
                                               + FlangeBottom_Thickness + ", " + Gasket_Id + ", '" + Gasket_Name + "', " + Gasket_Gap + ", "
                                               + Spacer_Id + ", '" + Spacer_Name + "', " + Spacer_Thickness + ",  " + ClampLength + ", '" + BoltDiameter_Id + "', "
                                               + "'" + Bolt_Thread + "', '" + Pitch_TPI_Title + "', " + PitchValue + ", " + BoltNumbers + ", " + LTFValue + ", "
                                               + BoltTool_RatioId + ", '" + BoltTool_Ratio + "', '" + Tool_Id + "', '" + Model_Number + "', "
                                               + ToolPressureArea + ", " + Material_Id + ", '" + Material_Value + "', " + BoltYieldValue + ", " + TensileStressAreaValue + ", "
                                               + BoltLengthValue + ", '" + BoltStressBase + "', " + GlobalMinorDiameterArea + ", " + T1ABoltStressValue + ", " + T1ABoltLoadValue + ", "
                                               + T1ABoltYieldValue + ", " + T1BBoltStressValue + ", " + T1BBoltLoadValue + ", " + T1BBoltYieldValue + ", "
                                               + T2RBoltStressValue + ", " + T2RBoltLoadValue + ", " + T2RBoltYieldValue + ", " + T3RBoltStressValue + ", " + T3RBoltLoadValue + ", " + T3RBoltYieldValue + ", " 
                                               + DetenBoltStressValue + ", " + DetenBoltLoadValue + ", "
                                               + DetenBoltYieldValue + ", " + Pass1AppliedPressure + ", " + Pass2AppliedPressure + ", " + Pass3AppliedPressure + ", " + CheckPass1AppliedPressure + ", "
                                               + TorqueValue + ", " + FrictionalCoefficientValue + ", " + "'1', '" + Pass2Bolt_SequenceTabValue_Label.Text + "', '" + Pass3Bolt_SequenceTabValue_Label.Text + "', '" + Pass4Bolt_SequenceTabValue_Label.Text + "',"
                                               + FirstPass_100 + ", " + FirstPass_50 + ", " + SecondPass_50 + ", " + Max_Detensioning + ", "
                                               + ResidualStressValue + ", '" + Comment_TextBox.Text + "', '" + IsDetensioned + "', " + CrossLoadingValue + ", " + DetensioningPerCent
                                               + ");";

                                    int AppResult = PopulateApplication_SQLiteCommand.ExecuteNonQuery();
                                }

                        }
                        else
                        {

                            PopulateApplication_SQLiteCommand.CommandText = "INSERT INTO ProjectReport (ProjectReportId, CustomerId, CustomerName, ProjectId, ProjectName, "
                                                                 + " ProjectReference, StartDateDay, StartDateMonth, StartDateYear, "
                                                                 + " EngineerId, EngineerName, SupervisorId, SupervisorName, Notes, SummaryNotes)"
                                                                 + " VALUES (0, 0, ' ', 1, ' ', ' ', "
                                                                 + Today.Day.ToString() + ", " + Today.Month.ToString() + ", " + Today.Year.ToString() + ", "
                                                                 + "1, ' ', " + "1,  ' ', ' ', ' ');";

                            int BlankProjectResult = PopulateApplication_SQLiteCommand.ExecuteNonQuery();

                            PopulateApplication_SQLiteCommand.CommandText = "";

                            PopulateApplication_SQLiteCommand.CommandText = "INSERT INTO ProjectDetailedReport (ProjectReportId, ApplicationId, UnitSystem, JointId, PrimaryStandardId,  "
                                               + "   Specification, PrimaryFlangeRatingId, FlangeRating, Flange1_TypeId,  "
                                               + "   Flange1_Abbreviation, Flange1_ClampLength, Flange2_TypeId, Flange2_Abbreviation,  "
                                               + "   Flange2_ClampLength, GasketId, Gasket, GasketGap,  "
                                               + "   SpacerId, Spacer, SpacerThickness, TotalClampLength, PrimaryFlangeBoltId,  "
                                               + "   BoltThread, Pitch_or_TPI, Pitch_TPI_Value, Number_of_Bolts, LTF, "
                                               + "   Bolt_to_ToolRatioId, Bolt_to_ToolRatio, ToolId, ModelNumber,  "
                                               + "   ToolPressureArea, MaterialId, Material, BoltYield, BoltTensileStressArea, "
                                               + "   BoltLength, BoltStressBase, BoltMinorDiameterArea, T1_APressureBoltStress, T1_APressureBoltLoad,  "
                                               + "   T1_APressureBoltYieldPC, T1_BPressureBoltStress, T1_BPressureBoltLoad, T1_BPressureBoltYieldPC,   "
                                               + "   T2_ResidualBoltStress, T2_ResidualBoltLoad, T2_ResidualBoltYieldPC, T3_ResidualBoltStress, T3_ResidualBoltLoad, T3_ResidualBoltYieldPC, "
                                               + "   DetensioningBoltStress, DetensioningBoltLoad,	"
                                               + "   DetensioningBoltYieldPC, TensionPressure_FirstPass, TensionPressure_SecondPass, TensionPressure_ThirdPass, CheckingPass,   "
                                               + "   Torque, Coefficient_Friction, Bolt01, Bolt02, Bolt03, Bolt04,"
                                               + "   FirstPass_100, FirstPass_50, SecondPass_50, Max_Detensioning, "
                                               + "   ResidualBoltStress, Comments, IsDetension, CrossLoadingPC, DetensioningPC"
                                               + ") "
                                               + " VALUES (" + ProjectId_Label.Text + ", " + (MaxAppID + 1).ToString() + ", '" + UnitSystem + "', '" + JointId_TextBox.Text + "', '" + Specification_ComboBox.SelectedValue.ToString() + "', "
                                               + "'" + SpecificationValue + "', '" + Rating_ComboBox.SelectedValue.ToString() + "', '" + Flange_Rating + "', " + Flange1Config_ComboBox.SelectedValue.ToString() + ", "
                                               + "'" + Flange1_AbbreviationValue + "', " + FlangeTop_Thickness + ", " + Flange2_Id + ", '" + Flange2_AbbreviationValue + "', "
                                               + FlangeBottom_Thickness + ", " + Gasket_Id + ", '" + Gasket_Name + "', " + Gasket_Gap + ", "
                                               + Spacer_Id + ", '" + Spacer_Name + "', " + Spacer_Thickness + ",  " + ClampLength + ", '" + BoltDiameter_Id + "', "
                                               + "'" + Bolt_Thread + "', '" + Pitch_TPI_Title + "', " + PitchValue + ", " + BoltNumbers + ", " + LTFValue + ", "
                                               + BoltTool_RatioId + ", '" + BoltTool_Ratio + "', '" + Tool_Id + "', '" + Model_Number + "', "
                                               + ToolPressureArea + ", " + Material_Id + ", '" + Material_Value + "', " + BoltYieldValue + ", " + TensileStressAreaValue + ", "
                                               + BoltLengthValue + ", '" + BoltStressBase + "', " + GlobalMinorDiameterArea + ", " + T1ABoltStressValue + ", " + T1ABoltLoadValue + ", "
                                               + T1ABoltYieldValue + ", " + T1BBoltStressValue + ", " + T1BBoltLoadValue + ", " + T1BBoltYieldValue + ", "
                                               + T2RBoltStressValue + ", " + T2RBoltLoadValue + ", " + T2RBoltYieldValue + ", " + T3RBoltStressValue + ", " + T3RBoltLoadValue + ", " + T3RBoltYieldValue + ", " 
                                               + DetenBoltStressValue + ", " + DetenBoltLoadValue + ", "
                                               + DetenBoltYieldValue + ", " + Pass1AppliedPressure + ", " + Pass2AppliedPressure + ", " + Pass3AppliedPressure + ", " + CheckPass1AppliedPressure + ", "
                                               + TorqueValue + ", " + FrictionalCoefficientValue + ", " + "'1', '" + Pass2Bolt_SequenceTabValue_Label.Text + "', '" + Pass3Bolt_SequenceTabValue_Label.Text + "', '" + Pass4Bolt_SequenceTabValue_Label.Text + "',"
                                               + FirstPass_100 + ", " + FirstPass_50 + ", " + SecondPass_50 + ", " + Max_Detensioning + ", "
                                               + ResidualStressValue + ", '" + Comment_TextBox.Text + "', '" + IsDetensioned + "', " + CrossLoadingValue + ", " + DetensioningPerCent
                                               + ");";

                            int BlankProjectDetailResult = PopulateApplication_SQLiteCommand.ExecuteNonQuery();
                        }
                    }
                    catch (Exception ProjectSave_Exception)
                    {
                        MessageBox.Show(ProjectSave_Exception.Message + " while saving new application.", "Fatal Error!!!");
                    }

                    if (CustomSpacer_TextBox.Text == "NULL")
                    {
                        CustomSpacer_TextBox.Text = "";
                    }
                    

                    try
                    {
                        PopulateApplication_SQLiteCommand.CommandText = "SELECT  PR.CustomerId, PR.ProjectReportId, ApplicationId, JointId, "
                                                                        + "   FlangeRating, Flange1_Abbreviation, Flange2_Abbreviation, "
                                                                        + "   (BoltThread || ' x ' ||	Pitch_TPI_Value) AS BoltSize, Number_of_Bolts, Material "
                                                                        + "   FROM ProjectReport   PR"
                                                                        + "   	INNER JOIN ProjectDetailedReport     PD 		ON		PD.ProjectReportId = PR.ProjectReportId "
                                                                        + "   	WHERE PR.ProjectReportId > 0  AND PR.ProjectReportId = " + ProjectId_Label.Text + ";";
                        DataSet AppGrid_DataSet = new DataSet();
                        SQLiteDataAdapter AppGrid_DataAdapter = new SQLiteDataAdapter(PopulateApplication_SQLiteCommand);
                        
                        
                        int AppGridResult = AppGrid_DataAdapter.Fill(AppGrid_DataSet);

                        // Data Grid Setup - Initially the Data Grid has not been binded to any data. Hence only first time you need to bind.
                        if (Project_DataGridView.DataSource != null)
                        {
                            Project_DataGridView.DataSource = null;
                        }
                        
                            Project_DataGridView.AutoGenerateColumns = false;
                            Project_DataGridView.ColumnCount = 10;
                            Project_DataGridView.Columns[0].Visible = false;
                            Project_DataGridView.Columns[0].Name = "Customer_Id";
                            Project_DataGridView.Columns[0].HeaderText = "Customer Id";
                            Project_DataGridView.Columns[0].DataPropertyName = "CustomerId";
                            Project_DataGridView.Columns[1].Visible = false;
                            Project_DataGridView.Columns[1].Name = "ProjectReport_Id";
                            Project_DataGridView.Columns[1].HeaderText = "Project Report Id";
                            Project_DataGridView.Columns[1].DataPropertyName = "ProjectReportId";
                            Project_DataGridView.Columns[2].Visible = false;
                            Project_DataGridView.Columns[2].Name = "Application_Id";
                            Project_DataGridView.Columns[2].HeaderText = "Application Id";
                            Project_DataGridView.Columns[2].DataPropertyName = "ApplicationId";
                            Project_DataGridView.Columns[3].Visible = true;
                            Project_DataGridView.Columns[3].Name = "Joint_Id";
                            Project_DataGridView.Columns[3].HeaderText = "ID";
                            Project_DataGridView.Columns[3].DataPropertyName = "JointId";
                            Project_DataGridView.Columns[3].Width = 120;
                            Project_DataGridView.Columns[4].Visible = true;
                            Project_DataGridView.Columns[4].Name = "Flange_Rating";
                            Project_DataGridView.Columns[4].HeaderText = "Rating";
                            Project_DataGridView.Columns[4].DataPropertyName = "FlangeRating";
                            Project_DataGridView.Columns[4].Width = 180;
                            Project_DataGridView.Columns[5].Visible = true;
                            Project_DataGridView.Columns[5].Name = "Flange1";
                            Project_DataGridView.Columns[5].HeaderText = "Flange 1";
                            Project_DataGridView.Columns[5].DataPropertyName = "Flange1_Abbreviation";
                            Project_DataGridView.Columns[6].Visible = true;
                            Project_DataGridView.Columns[6].Name = "Flange2";
                            Project_DataGridView.Columns[6].HeaderText = "Flange 2";
                            Project_DataGridView.Columns[6].DataPropertyName = "Flange2_Abbreviation";
                            Project_DataGridView.Columns[7].Visible = true;
                            Project_DataGridView.Columns[7].Name = "Bolt_Diameter";
                            Project_DataGridView.Columns[7].HeaderText = "Bolt Size";
                            Project_DataGridView.Columns[7].DataPropertyName = "BoltSize";
                            Project_DataGridView.Columns[8].Visible = true;
                            Project_DataGridView.Columns[8].Name = "Bolts";
                            Project_DataGridView.Columns[8].HeaderText = "Bolts";
                            Project_DataGridView.Columns[8].DataPropertyName = "Number_of_Bolts";
                            Project_DataGridView.Columns[9].Visible = true;
                            Project_DataGridView.Columns[9].Name = "Bolt_Material";
                            Project_DataGridView.Columns[9].HeaderText = "Bolt Material";
                            Project_DataGridView.Columns[9].DataPropertyName = "Material";
                            Project_DataGridView.Columns[9].Width = 160;
                            Project_DataGridView.DataSource = AppGrid_DataSet.Tables[0];
                        
                    }
                    catch (Exception AppRetrieval_Exception)
                    {
                        MessageBox.Show(AppRetrieval_Exception.Message + " while populating project data grid of new application.", "Error!");
                    }
                }
                else
                {
                    MessageBox.Show(" Residual bolt stress value is wrong. Correct it. Only numerals are allowed.");
                }
            }
        }

        private void BoltTool_ComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (!IsInitial)
            {
                if (!IsSetting)
                {
                        
                        // Bolt to tool proportion i.e. all bolts are tensioned simultaneously or not including tool not used at all.
                        OperateTool();
                        // Graph Tab
                        DrawGraph();
                }
            }
  
        }

        protected void SetApplication()
        {
            IsSetting = true;
            string BoltDiameter = string.Empty;
            decimal InfoPanel_ClampLength = 0M;
            decimal InfoPanel_ToolPressureArea = 0M;
            decimal InfoPanel_BoltYield = 0M;
            decimal InfoPanel_TensileStressArea = 0M;
            decimal InfoPanel_BoltLength = 0M;

            SQLiteConnection SetApp_Connection = new SQLiteConnection("Data Source=" + BasePath + "\\WizBolt.db;New=False;");
            SQLiteCommand SetApp_Command = new SQLiteCommand(SetApp_Connection);
            SetApp_Connection.Open();

            SetApp_Command.CommandText = " SELECT ProjectReportId, ApplicationId, UnitSystem, JointId, PrimaryStandardId,  "
                                           + "   Specification, PrimaryFlangeRatingId, FlangeRating, Flange1_TypeId,  "
                                           + "   Flange1_Abbreviation, Flange1_ClampLength, Flange2_TypeId, Flange2_Abbreviation,  "
                                           + "   Flange2_ClampLength, GasketId, Gasket, GasketGap,  "
                                           + "   SpacerId, Spacer, SpacerThickness, TotalClampLength, PrimaryFlangeBoltId,  "
                                           + "   BoltThread, Pitch_or_TPI, Pitch_TPI_Value, Number_of_Bolts, LTF, "
                                           + "   Bolt_to_ToolRatioId, Bolt_to_ToolRatio, ToolId, ModelNumber,  "
                                           + "   ToolPressureArea, MaterialId, Material, BoltYield, BoltTensileStressArea, "
                                           + "   BoltLength, BoltStressBase, BoltMinorDiameterArea, T1_APressureBoltStress, T1_APressureBoltLoad,  "
                                           + "   T1_APressureBoltYieldPC, T1_BPressureBoltStress, T1_BPressureBoltLoad, T1_BPressureBoltYieldPC,   "
                                           + "   T2_ResidualBoltStress, T2_ResidualBoltLoad, T2_ResidualBoltYieldPC, T3_ResidualBoltStress, T3_ResidualBoltLoad, T3_ResidualBoltYieldPC,"
                                           + "   DetensioningBoltStress, DetensioningBoltLoad,	"
                                           + "   DetensioningBoltYieldPC, TensionPressure_FirstPass, TensionPressure_SecondPass, TensionPressure_ThirdPass, CheckingPass,   "
                                           + "   Torque, Coefficient_Friction, Bolt01, Bolt02, Bolt03, Bolt04, "
                                           + "   FirstPass_100, FirstPass_50, SecondPass_50, Max_Detensioning, "
                                           + "   ResidualBoltStress, Comments, IsDetension, CrossLoadingPC, DetensioningPC   "
                                           + "   FROM ProjectDetailedReport "
                                           + "           WHERE        ProjectReportId =  " + ProjectId_Label.Text.Trim()
                                           + "                  AND   ApplicationId   =  " + AppId_Label.Text.Trim();
            DataSet SelectedApp_DataSet = new DataSet();
            SQLiteDataAdapter SelectedApp_DataAdapter = new SQLiteDataAdapter(SetApp_Command);
            int SelectedAppResult = SelectedApp_DataAdapter.Fill(SelectedApp_DataSet);

            // Assign
            if (SelectedAppResult > 0)
            {
                
                if (SelectedApp_DataSet.Tables[0].Rows[0]["UnitSystem"].ToString().Trim() == "SAE")
                {
                    UnitSystem_Button.Text = "SAE";
                    UnitSystem_Button.BorderColor = Color.SteelBlue;
                    UnitSystem_Button.ColorFillBlend.iColor[0] = Color.AliceBlue;
                    UnitSystem_Button.ColorFillBlend.iColor[1] = Color.RoyalBlue;
                    UnitSystem_Button.ColorFillBlend.iColor[2] = Color.Navy;
                    Clamp1Unit_Label.Text = "in";
                    GasketGapUnit_Label.Text = "in";
                    Clamp2Unit_Label.Text = "in";
                    CustomSpacerUnit_Label.Text = "in";
                    ResidualBoltStressUnit_Label.Text = "psi";
                }
                else
                {
                    UnitSystem_Button.Text = "S. I. Units";
                    UnitSystem_Button.BorderColor = Color.ForestGreen;
                    UnitSystem_Button.ColorFillBlend.iColor[0] = Color.GreenYellow;
                    UnitSystem_Button.ColorFillBlend.iColor[1] = Color.ForestGreen;
                    UnitSystem_Button.ColorFillBlend.iColor[2] = Color.DarkGreen;
                    Clamp1Unit_Label.Text = "mm";
                    GasketGapUnit_Label.Text = "mm";
                    Clamp2Unit_Label.Text = "mm";
                    CustomSpacerUnit_Label.Text = "mm";
                    ResidualBoltStressUnit_Label.Text = "MPa";
                }
                JointId_TextBox.Text = SelectedApp_DataSet.Tables[0].Rows[0]["JointId"].ToString();
                Comment_TextBox.Text = SelectedApp_DataSet.Tables[0].Rows[0]["Comments"].ToString();
                Specification_ComboBox.SelectedValue = SelectedApp_DataSet.Tables[0].Rows[0]["PrimaryStandardId"].ToString();
                Specification_ComboBox.SelectedItem = SelectedApp_DataSet.Tables[0].Rows[0]["Specification"].ToString();
                Rating_ComboBox.SelectedValue = SelectedApp_DataSet.Tables[0].Rows[0]["PrimaryFlangeRatingId"].ToString();
                Rating_ComboBox.SelectedItem = SelectedApp_DataSet.Tables[0].Rows[0]["FlangeRating"].ToString();
                BoltMaterial_ComboBox.SelectedValue = SelectedApp_DataSet.Tables[0].Rows[0]["MaterialId"].ToString();
                BoltMaterial_ComboBox.SelectedItem = SelectedApp_DataSet.Tables[0].Rows[0]["Material"].ToString();
                
                Flange1Config_ComboBox.SelectedValue = SelectedApp_DataSet.Tables[0].Rows[0]["Flange1_TypeId"].ToString();
                Flange1Config_ComboBox.SelectedItem = SelectedApp_DataSet.Tables[0].Rows[0]["Flange1_Abbreviation"].ToString();
                Flange2Config_ComboBox.SelectedValue = SelectedApp_DataSet.Tables[0].Rows[0]["Flange2_TypeId"].ToString();
                Flange2Config_ComboBox.SelectedItem = SelectedApp_DataSet.Tables[0].Rows[0]["Flange2_Abbreviation"].ToString();
                if (SelectedApp_DataSet.Tables[0].Rows[0]["Gasket"].ToString() != null)
                {
                    Gasket_ComboBox.SelectedValue = SelectedApp_DataSet.Tables[0].Rows[0]["GasketId"].ToString();
                    Gasket_ComboBox.SelectedItem = SelectedApp_DataSet.Tables[0].Rows[0]["Gasket"].ToString();
                }
                if (SelectedApp_DataSet.Tables[0].Rows[0]["SpacerId"].ToString() != null)
                {
                    Spacer_ComboBox.SelectedValue = SelectedApp_DataSet.Tables[0].Rows[0]["SpacerId"].ToString();
                    Spacer_ComboBox.SelectedItem = SelectedApp_DataSet.Tables[0].Rows[0]["Spacer"].ToString();
                }
                if (UnitSystem_Button.Text == "SAE")
                {
                    Clamp1Unit_Label.Text = "in";
                    GasketGapUnit_Label.Text = "in";
                    Clamp2Unit_Label.Text = "in";
                    CustomSpacerUnit_Label.Text = "in";
                    ResidualBoltStressUnit_Label.Text = "psi";
                    if ((SelectedApp_DataSet.Tables[0].Rows[0]["Flange1_ClampLength"].ToString().Trim() != "") && (SelectedApp_DataSet.Tables[0].Rows[0]["Flange1_ClampLength"].ToString().Trim().Length > 0))
                    {
                        Clamp1_TextBox.Text = Math.Round((Convert.ToDecimal(SelectedApp_DataSet.Tables[0].Rows[0]["Flange1_ClampLength"].ToString()) / 25.4M), 4, MidpointRounding.AwayFromZero).ToString();
                    }
                    else
                    {
                        Clamp1_TextBox.Text = "";
                    }
                    if ((SelectedApp_DataSet.Tables[0].Rows[0]["Flange2_ClampLength"].ToString().Trim() != "") && (SelectedApp_DataSet.Tables[0].Rows[0]["Flange2_ClampLength"].ToString().Trim().Length > 0))
                    {
                        Clamp2_TextBox.Text = Math.Round((Convert.ToDecimal(SelectedApp_DataSet.Tables[0].Rows[0]["Flange2_ClampLength"].ToString()) / 25.4M), 4, MidpointRounding.AwayFromZero).ToString();
                    }
                    else
                    {
                        Clamp2_TextBox.Text = "";
                    }

                    if ((SelectedApp_DataSet.Tables[0].Rows[0]["GasketGap"].ToString().Trim().Length > 0) && (SelectedApp_DataSet.Tables[0].Rows[0]["GasketGap"].ToString().Trim() != ""))
                    {
                        try
                        {
                             decimal GapTest = Convert.ToDecimal(SelectedApp_DataSet.Tables[0].Rows[0]["GasketGap"].ToString());
                             GasketGap_TextBox.Text = Math.Round((Convert.ToDecimal(SelectedApp_DataSet.Tables[0].Rows[0]["GasketGap"].ToString()) / 25.4M), 4, MidpointRounding.AwayFromZero).ToString();
                        }
                        catch (Exception Gap_Exception)
                        {
                            MessageBox.Show(Gap_Exception.Message + " Gasket thickness that is included in clamp is found to be wrong, while setting application.");
                            GasketGap_TextBox.Text = "";
                        }
                    }
                    else
                    {
                        GasketGap_TextBox.Text = "";
                    }
                    if ((SelectedApp_DataSet.Tables[0].Rows[0]["SpacerThickness"].ToString().Trim().Length > 0) && (SelectedApp_DataSet.Tables[0].Rows[0]["SpacerThickness"].ToString().Trim() != ""))
                    {
                        try
                        {
                            decimal SpacerTest = Convert.ToDecimal(SelectedApp_DataSet.Tables[0].Rows[0]["SpacerThickness"].ToString());

                            CustomSpacer_TextBox.Text = Math.Round((Convert.ToDecimal(SelectedApp_DataSet.Tables[0].Rows[0]["SpacerThickness"].ToString()) / 25.4M), 4, MidpointRounding.AwayFromZero).ToString();

                        }
                        catch (Exception Spacer_Exception)
                        {
                            MessageBox.Show(Spacer_Exception.Message + " Spacer thickness found to be wrong while setting application.");
                            CustomSpacer_TextBox.Text = "";
                        }
                        
                    }
                    else
                    {
                        CustomSpacer_TextBox.Text = "";
                    }
                    
                }
                else
                {
                    Clamp1Unit_Label.Text = "mm";
                    GasketGapUnit_Label.Text = "mm";
                    Clamp2Unit_Label.Text = "mm";
                    CustomSpacerUnit_Label.Text = "mm";
                    ResidualBoltStressUnit_Label.Text = "MPa";
                    Clamp1_TextBox.Text = SelectedApp_DataSet.Tables[0].Rows[0]["Flange1_ClampLength"].ToString();
                    Clamp2_TextBox.Text = SelectedApp_DataSet.Tables[0].Rows[0]["Flange2_ClampLength"].ToString();
                    if (Convert.ToDecimal(SelectedApp_DataSet.Tables[0].Rows[0]["GasketGap"].ToString()) > 0)
                    {
                        GasketGap_TextBox.Text = SelectedApp_DataSet.Tables[0].Rows[0]["GasketGap"].ToString();
                    }
                    else
                    {
                        GasketGap_TextBox.Text = "";
                    }
                    if (SelectedApp_DataSet.Tables[0].Rows[0]["SpacerThickness"].ToString() != "")
                    {
                        CustomSpacer_TextBox.Text = SelectedApp_DataSet.Tables[0].Rows[0]["SpacerThickness"].ToString();
                    }
                    else
                    {
                        CustomSpacer_TextBox.Text = "";
                    }
                }
                
                try
                {
                    string FlangeType = SelectedApp_DataSet.Tables[0].Rows[0]["Flange1_TypeId"].ToString();
                    string FlangePrimaryId = SelectedApp_DataSet.Tables[0].Rows[0]["PrimaryFlangeRatingId"].ToString();
                    string[] FlangeId = FlangePrimaryId.Split('~');
                    SetApp_Command.CommandText = "SELECT    (F.StandardId  ||  '~'  ||  F.StandardAuxiliaryId  ||  '~'  ||  F.FlangeRatingId  ||  '~'  ||  F.FlangeId || '~' || FB.FlangeBoltId)   AS PrimaryFlangeBoltId, FB.BoltThread, FB.TPI_Pitch, FB.BoltLength   "
                                                      + "   FROM  FlangeBolts		FB  "
                                                      + "   INNER JOIN	Flanges					F			ON					F.StandardId					=		FB.StandardId   "
                                                      + "                               AND		F.StandardAuxiliaryId		=		FB.StandardAuxiliaryId  "
                                                      + "                               AND		F.FlangeRatingId				=		FB.FlangeRatingId   "
                                                      + "								AND		F.FlangeId						=		FB.FlangeId         "
                                                      + " WHERE					F.FlangeTypeId			=    " + FlangeType
                                                      + "               AND		F.StandardId			=    " + FlangeId[0]
                                                      + "               AND		F.StandardAuxiliaryId  	=    " + FlangeId[1]
                                                      + "               AND  	F.FlangeRatingId 		=    " + FlangeId[2];

                    DataSet FlangeBolts_DataSet = new DataSet();
                    SQLiteDataAdapter FlangeBolts_DataAdapter = new SQLiteDataAdapter(SetApp_Command);
                    int FlangeBoltsResult = FlangeBolts_DataAdapter.Fill(FlangeBolts_DataSet);
                    if ((FlangeBoltsResult > 0) && (FlangeBolts_DataSet.Tables[0].Rows.Count > 0))
                    {
                        BoltThread_ComboBox.DataSource = FlangeBolts_DataSet.Tables[0];
                        BoltThread_ComboBox.DisplayMember = "BoltThread";
                        BoltThread_ComboBox.ValueMember = "PrimaryFlangeBoltId";
                        BoltThread_ComboBox.SelectedValue = SelectedApp_DataSet.Tables[0].Rows[0]["PrimaryFlangeBoltId"];
                        BoltThread_ComboBox.SelectedItem = SelectedApp_DataSet.Tables[0].Rows[0]["BoltThread"];
                       
                    }
            }
            catch (Exception FlangeBolts_Exception)
            {
                MessageBox.Show(FlangeBolts_Exception.Message + " while setting flange bolts.");
            }
                if (SelectedApp_DataSet.Tables[0].Rows[0]["ModelNumber"].ToString().Trim().Length > 0)
                {
                   // string[] ToolType = SelectedApp_DataSet.Tables[0].Rows[0]["ToolId"].ToString().Split('~');
                   //// ToolSeriesId_Label.Text = ToolType[0];
                   // SetApp_Command.CommandText = "SELECT TensionerToolSeriesId,  (COALESCE(AppliedField, ' ') || ' ' || COALESCE(Series, '')) AS ToolSeries "
                   //                              + " FROM TensionerTools_Series WHERE TensionerToolSeriesId = " + ToolType[0];

                   // DataSet ToolSeries_DataSet = new DataSet();
                   // SQLiteDataAdapter ToolSeries_DataAdapter = new SQLiteDataAdapter(SetApp_Command);
                   // try
                   // {
                   //     int ToolSeriesResult = ToolSeries_DataAdapter.Fill(ToolSeries_DataSet);
                   //     //if (ToolSeriesResult > 0)
                   //     //{
                   //     //    TensionerToolRange_Label.Text = ToolSeries_DataSet.Tables[0].Rows[0]["ToolSeries"].ToString();
                   //     //}
                   // }
                   // catch (Exception ToolSeries_Exception)
                   // {
                   //     MessageBox.Show(ToolSeries_Exception.Message + " while setting the application.");
                   // }

                    // Populate the Tool Drop down list
                    string SelectedBoltThread = BoltThread_ComboBox.SelectedValue.ToString();
                    string SelectedTool_List_of_Ids = string.Empty;
                    string[] Tool_Identities = ToolSeriesId_Label.Text.Trim().Split('&');
                    string[] ToolParams = SelectedBoltThread.Split('~');
                    string Auxiliary_and_Standard_Id = (Specification_ComboBox.SelectedValue).ToString();  // Standard Id and Auxiliary Id in the form of 1~1
                    string[] Separated_Id = Auxiliary_and_Standard_Id.Split('~');               // Standard Id and Auxiliary Ids split in individual
                    for (int i = 0; i < Tool_Identities.Count(); i++)
                    {
                        if (i == 0)
                        {
                            SelectedTool_List_of_Ids = Tool_Identities[i];
                        }
                        else
                        {
                            SelectedTool_List_of_Ids = SelectedTool_List_of_Ids + ", " + Tool_Identities[i];
                        }
                    }
                    if (Convert.ToInt32(Separated_Id[0]) == 0)
                    {
                        SetApp_Command.CommandText = "   SELECT DISTINCT (F.TensionerToolSeriesId || '~' || F.TensionerToolId) AS ToolId,  T.ModelNumber, T.HydraulicArea "
                                                + "  FROM  FlangeBoltTools		F	"
                                                + "                 INNER JOIN   TensionerTools  T    ON       T.TensionerToolSeriesId   =  F.TensionerToolSeriesId "
                                                + "                                                        AND T.TensionerToolId         =  F.TensionerToolId  "
                                                + "                  WHERE		F.StandardId =   " + ToolParams[0]
                                                + "                         AND F.StandardAuxiliaryId =  " + ToolParams[1]
                                                + "                         AND F.FlangeRatingId   =  " + ToolParams[2]
                                                + "                         AND F.FlangeId         =  " + ToolParams[3]
                                                + "                         AND F.FlangeBoltId     =  " + ToolParams[4]
                                                + "                         AND F.TensionerToolSeriesId IN (5, " + SelectedTool_List_of_Ids + ");";
                    }
                    else
                    {
                        SetApp_Command.CommandText = "   SELECT (F.TensionerToolSeriesId || '~' || F.TensionerToolId) AS ToolId,  T.ModelNumber, T.HydraulicArea "
                                        + "  FROM  FlangeBoltTools		F	"
                                        + "                 INNER JOIN   TensionerTools  T    ON       T.TensionerToolSeriesId   =  F.TensionerToolSeriesId "
                                        + "                                                        AND T.TensionerToolId         =  F.TensionerToolId  "
                                        + "                  WHERE		F.StandardId =   " + ToolParams[0]
                                        + "                         AND F.StandardAuxiliaryId =  " + ToolParams[1]
                                        + "                         AND F.FlangeRatingId   =  " + ToolParams[2]
                                        + "                         AND F.FlangeId         =  " + ToolParams[3]
                                        + "                         AND F.FlangeBoltId     =  " + ToolParams[4]
                                        + "                          AND F.TensionerToolSeriesId IN (" + SelectedTool_List_of_Ids + ");";

                    }

                    try
                    {
                        DataSet Tools_DataSet = new DataSet();
                        SQLiteDataAdapter Tools_DataAdapter = new SQLiteDataAdapter(SetApp_Command);
                        int ToolResult = Tools_DataAdapter.Fill(Tools_DataSet);
                        if ((ToolResult > 0) && (Tools_DataSet.Tables[0].Rows.Count > 0))
                        {
                            TensionerTool_ComboBox.DataSource = Tools_DataSet.Tables[0];
                            TensionerTool_ComboBox.DisplayMember = "ModelNumber";
                            TensionerTool_ComboBox.ValueMember = "ToolId";
                        }
                        else
                        {
                            TensionerTool_ComboBox.DataSource = null;
                            TensionerTool_ComboBox.DisplayMember = null;
                            TensionerTool_ComboBox.ValueMember = null;
                        }
                    }
                    catch (Exception Tools_Exception)
                    {
                        MessageBox.Show(Tools_Exception.Message + " while populating tools during bolt sizing event.");
                    }


                    // Set the values
                    TensionerTool_ComboBox.SelectedValue = SelectedApp_DataSet.Tables[0].Rows[0]["ToolId"].ToString();
                    TensionerTool_ComboBox.SelectedItem = SelectedApp_DataSet.Tables[0].Rows[0]["ModelNumber"].ToString();
                }
                else
                {
                    TensionerTool_ComboBox.Items.Clear();
                } 
                Pitch_Label.Text = SelectedApp_DataSet.Tables[0].Rows[0]["Pitch_or_TPI"].ToString();
                Pitch_TextBox.Text = SelectedApp_DataSet.Tables[0].Rows[0]["Pitch_TPI_Value"].ToString();
                NumberOfBolts_TextBox.Text = SelectedApp_DataSet.Tables[0].Rows[0]["Number_of_Bolts"].ToString();
                BoltTool_ComboBox.SelectedValue = SelectedApp_DataSet.Tables[0].Rows[0]["Bolt_to_ToolRatioId"].ToString();
                BoltTool_ComboBox.SelectedItem = SelectedApp_DataSet.Tables[0].Rows[0]["Bolt_to_ToolRatio"].ToString();
                ResidualBoltStress_TextBox.Text = SelectedApp_DataSet.Tables[0].Rows[0]["ResidualBoltStress"].ToString();
                CrossLoading_TextBox.Text = SelectedApp_DataSet.Tables[0].Rows[0]["CrossLoadingPC"].ToString();
                Detensioning_TextBox.Text = SelectedApp_DataSet.Tables[0].Rows[0]["DetensioningPC"].ToString();
                if (SelectedApp_DataSet.Tables[0].Rows[0]["IsDetension"].ToString().Trim() == "Yes")
                {
                    DeTension_CheckBox.Checked = true;
                }
                else
                {
                    DeTension_CheckBox.Checked = false;
                }
                if (SelectedApp_DataSet.Tables[0].Rows[0]["Bolt_to_ToolRatio"].ToString().Trim() == "100%")
                {
                    PressureA_In_Label.Text = "";
                    PressureA_SI_Label.Text = "";
                    if ((SelectedApp_DataSet.Tables[0].Rows[0]["FirstPass_100"].ToString().Trim() != "") && (SelectedApp_DataSet.Tables[0].Rows[0]["FirstPass_100"].ToString().Trim().Length > 0))
                    {
                        try
                        {
                            decimal PassPressureTest = Convert.ToDecimal(SelectedApp_DataSet.Tables[0].Rows[0]["FirstPass_100"].ToString());
                            PressureB_In_Label.Text = Math.Round((Convert.ToDecimal(SelectedApp_DataSet.Tables[0].Rows[0]["FirstPass_100"].ToString()) * 145.037738007M), 0, MidpointRounding.AwayFromZero).ToString() + " psi";
                            PressureB_SI_Label.Text = SelectedApp_DataSet.Tables[0].Rows[0]["FirstPass_100"].ToString() + " MPa";
                            Detensioning_In_Label.Text = Math.Round((Convert.ToDecimal(SelectedApp_DataSet.Tables[0].Rows[0]["FirstPass_100"].ToString()) * 145.037738007M), 0, MidpointRounding.AwayFromZero).ToString() + " psi";
                            Detensioning_SI_Label.Text = SelectedApp_DataSet.Tables[0].Rows[0]["FirstPass_100"].ToString() + " MPa";
                        }
                        catch (Exception FPressure_Exception)
                        {
                            MessageBox.Show(FPressure_Exception.Message + " Tightening pressure applied to all bolts found wrong while setting application.");
                            PressureB_In_Label.Text = "";
                            PressureB_SI_Label.Text = "";
                            Detensioning_In_Label.Text = "";
                            Detensioning_SI_Label.Text = "";
                        }
                    }
                    else
                    {
                        PressureB_In_Label.Text = "";
                        PressureB_SI_Label.Text = "";
                        Detensioning_In_Label.Text = "";
                        Detensioning_SI_Label.Text = "";
                    }

                }
                else
                {
                    if ((SelectedApp_DataSet.Tables[0].Rows[0]["FirstPass_50"].ToString().Trim() != "") && (SelectedApp_DataSet.Tables[0].Rows[0]["FirstPass_50"].ToString().Trim().Length > 0))
                    {
                        try
                        {
                            decimal FPassTest = Convert.ToDecimal(SelectedApp_DataSet.Tables[0].Rows[0]["FirstPass_50"].ToString());
                            PressureA_In_Label.Text = Math.Round((Convert.ToDecimal(SelectedApp_DataSet.Tables[0].Rows[0]["FirstPass_50"].ToString()) * 145.037738007M), 0, MidpointRounding.AwayFromZero).ToString() + " psi";
                            PressureA_SI_Label.Text = SelectedApp_DataSet.Tables[0].Rows[0]["FirstPass_50"].ToString() + " MPa";
                        }
                        catch (Exception FPass_Exception)
                        {
                            MessageBox.Show(FPass_Exception.Message + " Half number of bolts first pass pressure found wrong while setting application.");
                            PressureA_In_Label.Text = "";
                            PressureA_SI_Label.Text = "";
                        }
                    }
                    else
                    {
                        PressureA_In_Label.Text = "";
                        PressureA_SI_Label.Text = "";
                    }
                    if ((SelectedApp_DataSet.Tables[0].Rows[0]["SecondPass_50"].ToString().Trim() != "") && (SelectedApp_DataSet.Tables[0].Rows[0]["SecondPass_50"].ToString().Trim().Length > 0))
                    {
                        try
                        {
                            decimal SPressureTest = Convert.ToDecimal(SelectedApp_DataSet.Tables[0].Rows[0]["SecondPass_50"].ToString());
                            PressureB_In_Label.Text = Math.Round((Convert.ToDecimal(SelectedApp_DataSet.Tables[0].Rows[0]["SecondPass_50"].ToString()) * 145.037738007M), 0, MidpointRounding.AwayFromZero).ToString() + " psi";
                            PressureB_SI_Label.Text = SelectedApp_DataSet.Tables[0].Rows[0]["SecondPass_50"].ToString() + " MPa";
                            Detensioning_In_Label.Text = Math.Round((Convert.ToDecimal(SelectedApp_DataSet.Tables[0].Rows[0]["SecondPass_50"].ToString()) * 145.037738007M), 0, MidpointRounding.AwayFromZero).ToString() + " psi";
                            Detensioning_SI_Label.Text = SelectedApp_DataSet.Tables[0].Rows[0]["SecondPass_50"].ToString() + " MPa";
                        }
                        catch(Exception SPressure_Exception)
                        {
                            MessageBox.Show(SPressure_Exception.Message + " Half number of bolts tightening pressure value in second pass found to be wrong.");
                            PressureB_In_Label.Text = "";
                            PressureB_SI_Label.Text = "";
                            Detensioning_In_Label.Text = "";
                            Detensioning_SI_Label.Text = "";
                        }
                    }
                    else
                    {
                        PressureB_In_Label.Text = "";
                        PressureB_SI_Label.Text = "";
                        Detensioning_In_Label.Text = "";
                        Detensioning_SI_Label.Text = "";
                    }
                }
                Torque_SI_Label.Text = SelectedApp_DataSet.Tables[0].Rows[0]["Torque"].ToString() + " N-m";
                if ((SelectedApp_DataSet.Tables[0].Rows[0]["Torque"].ToString().Trim() != "") && (SelectedApp_DataSet.Tables[0].Rows[0]["Torque"].ToString().Trim().Length > 0))
                {
                    try
                    {
                        decimal TorqueTest = Convert.ToDecimal(SelectedApp_DataSet.Tables[0].Rows[0]["Torque"].ToString());
                        Torque_In_Label.Text = Math.Round((Convert.ToDecimal(SelectedApp_DataSet.Tables[0].Rows[0]["Torque"].ToString()) / 1.35581794833M), 0, MidpointRounding.AwayFromZero).ToString() + " ft-lbs";
                    }
                    catch(Exception Torque_Exception)
                    {
                        MessageBox.Show(Torque_Exception.Message + " Torque is found wrong while setting the application.");
                        Torque_In_Label.Text = "";
                    }
                }
                else
                {
                    Torque_In_Label.Text = "";
                }
                TorqueTab_SI_Label.Text = Torque_SI_Label.Text;
                TorqueValueTab_In_Label.Text = Torque_In_Label.Text;

                Torque_SI_Sequence_Label.Text = Torque_SI_Label.Text;
                Torque_In_Sequence_Label.Text = Torque_In_Label.Text;

                // Bolt Stress Tab
                if ((SelectedApp_DataSet.Tables[0].Rows[0]["T1_APressureBoltStress"].ToString().Trim() != "") && (SelectedApp_DataSet.Tables[0].Rows[0]["T1_APressureBoltStress"].ToString().Trim().Length > 0))
                {
                    try
                    {
                        decimal T1_APTest = Convert.ToDecimal(SelectedApp_DataSet.Tables[0].Rows[0]["T1_APressureBoltStress"].ToString().Trim());
                        T1ABoltStress_In_Label.Text = Math.Round(T1_APTest * 145.037738007M, 2, MidpointRounding.AwayFromZero).ToString();
                    }
                    catch (Exception T1_APStress_Exception)
                    {
                        MessageBox.Show(T1_APStress_Exception.Message + " while setting Bolt Stress Tab.");
                        T1ABoltStress_In_Label.Text = "";
                    }
                }
                else
                {
                    T1ABoltStress_In_Label.Text = "";
                }
                T1ABoltStress_SI_Label.Text = SelectedApp_DataSet.Tables[0].Rows[0]["T1_APressureBoltStress"].ToString().Trim();

                if ((SelectedApp_DataSet.Tables[0].Rows[0]["T1_APressureBoltLoad"].ToString().Trim() != "") && (SelectedApp_DataSet.Tables[0].Rows[0]["T1_APressureBoltLoad"].ToString().Trim().Length > 0))
                {
                    try
                    {
                        decimal T1_APLTest = Convert.ToDecimal(SelectedApp_DataSet.Tables[0].Rows[0]["T1_APressureBoltLoad"].ToString().Trim());
                        T1ABoltLoad_In_Label.Text = Math.Round(T1_APLTest * 145.037738007M, 2, MidpointRounding.AwayFromZero).ToString();
                    }
                    catch (Exception T1_APStress_Exception)
                    {
                        MessageBox.Show(T1_APStress_Exception.Message + " while setting Bolt Stress Tab.");
                        T1ABoltLoad_In_Label.Text = "";
                    }
                }
                else
                {
                    T1ABoltLoad_In_Label.Text = "";
                }
                T1ABoltLoad_SI_Label.Text = SelectedApp_DataSet.Tables[0].Rows[0]["T1_APressureBoltLoad"].ToString().Trim();
                T1ABoltYield_SI_Label.Text = SelectedApp_DataSet.Tables[0].Rows[0]["T1_APressureBoltYieldPC"].ToString().Trim();

                // Row 2
                if ((SelectedApp_DataSet.Tables[0].Rows[0]["T1_BPressureBoltStress"].ToString().Trim() != "") && (SelectedApp_DataSet.Tables[0].Rows[0]["T1_BPressureBoltStress"].ToString().Trim().Length > 0))
                {
                    try
                    {
                        decimal T1_BPTest = Convert.ToDecimal(SelectedApp_DataSet.Tables[0].Rows[0]["T1_BPressureBoltStress"].ToString().Trim());
                        T1BBoltStress_In_Label.Text = Math.Round(T1_BPTest * 145.037738007M, 2, MidpointRounding.AwayFromZero).ToString();
                    }
                    catch (Exception T1_APStress_Exception)
                    {
                        MessageBox.Show(T1_APStress_Exception.Message + " while setting Bolt Stress Tab.");
                        T1BBoltStress_In_Label.Text = "";
                    }
                }
                else
                {
                    T1BBoltStress_In_Label.Text = "";
                }
                T1BBoltStress_SI_Label.Text = SelectedApp_DataSet.Tables[0].Rows[0]["T1_BPressureBoltStress"].ToString().Trim();

                if ((SelectedApp_DataSet.Tables[0].Rows[0]["T1_BPressureBoltLoad"].ToString().Trim() != "") && (SelectedApp_DataSet.Tables[0].Rows[0]["T1_BPressureBoltLoad"].ToString().Trim().Length > 0))
                {
                    try
                    {
                        decimal T1_BPLTest = Convert.ToDecimal(SelectedApp_DataSet.Tables[0].Rows[0]["T1_BPressureBoltLoad"].ToString().Trim());
                        T1BBoltLoad_In_Label.Text = Math.Round(T1_BPLTest * 145.037738007M, 2, MidpointRounding.AwayFromZero).ToString();
                    }
                    catch (Exception T1_APStress_Exception)
                    {
                        MessageBox.Show(T1_APStress_Exception.Message + " while setting Bolt Stress Tab.");
                        T1BBoltLoad_In_Label.Text = "";
                    }
                }
                else
                {
                    T1BBoltLoad_In_Label.Text = "";
                }
                T1BBoltLoad_SI_Label.Text = SelectedApp_DataSet.Tables[0].Rows[0]["T1_BPressureBoltLoad"].ToString().Trim();
                T1BBoltYield_Per_Label.Text = SelectedApp_DataSet.Tables[0].Rows[0]["T1_BPressureBoltYieldPC"].ToString().Trim();

                // Row 3
                if ((SelectedApp_DataSet.Tables[0].Rows[0]["T2_ResidualBoltStress"].ToString().Trim() != "") && (SelectedApp_DataSet.Tables[0].Rows[0]["T2_ResidualBoltStress"].ToString().Trim().Length > 0))
                {
                    try
                    {
                        decimal T2_RTest = Convert.ToDecimal(SelectedApp_DataSet.Tables[0].Rows[0]["T2_ResidualBoltStress"].ToString().Trim());
                        T2RBoltStress_In_Label.Text = Math.Round(T2_RTest * 145.037738007M, 2, MidpointRounding.AwayFromZero).ToString();
                    }
                    catch (Exception T1_APStress_Exception)
                    {
                        MessageBox.Show(T1_APStress_Exception.Message + " while setting Bolt Stress Tab.");
                        T2RBoltStress_In_Label.Text = "";
                    }
                }
                else
                {
                    T2RBoltStress_In_Label.Text = "";
                }
                T2RBoltStress_SI_Label.Text = SelectedApp_DataSet.Tables[0].Rows[0]["T2_ResidualBoltStress"].ToString().Trim();

                if ((SelectedApp_DataSet.Tables[0].Rows[0]["T2_ResidualBoltLoad"].ToString().Trim() != "") && (SelectedApp_DataSet.Tables[0].Rows[0]["T2_ResidualBoltLoad"].ToString().Trim().Length > 0))
                {
                    try
                    {
                        decimal T1_BPLTest = Convert.ToDecimal(SelectedApp_DataSet.Tables[0].Rows[0]["T2_ResidualBoltLoad"].ToString().Trim());
                        T2RBoltLoad_In_Label.Text = Math.Round(T1_BPLTest * 145.037738007M, 2, MidpointRounding.AwayFromZero).ToString();
                    }
                    catch (Exception T1_APStress_Exception)
                    {
                        MessageBox.Show(T1_APStress_Exception.Message + " while setting Bolt Stress Tab.");
                        T2RBoltLoad_In_Label.Text = "";
                    }
                }
                else
                {
                    T2RBoltLoad_In_Label.Text = "";
                }
                T2RBoltLoad_SI_Label.Text = SelectedApp_DataSet.Tables[0].Rows[0]["T2_ResidualBoltLoad"].ToString().Trim();
                T2RBoltYield_Per_Label.Text = SelectedApp_DataSet.Tables[0].Rows[0]["T2_ResidualBoltYieldPC"].ToString().Trim();

                // Row 4
                if ((SelectedApp_DataSet.Tables[0].Rows[0]["T3_ResidualBoltStress"].ToString().Trim() != "") && (SelectedApp_DataSet.Tables[0].Rows[0]["T3_ResidualBoltStress"].ToString().Trim().Length > 0))
                {
                    try
                    {
                        decimal T2_RTest = Convert.ToDecimal(SelectedApp_DataSet.Tables[0].Rows[0]["T3_ResidualBoltStress"].ToString().Trim());
                        T3RBoltStress_In_Label.Text = Math.Round(T2_RTest * 145.037738007M, 2, MidpointRounding.AwayFromZero).ToString();
                    }
                    catch (Exception T1_APStress_Exception)
                    {
                        MessageBox.Show(T1_APStress_Exception.Message + " while setting Bolt Stress Tab.");
                        T3RBoltStress_In_Label.Text = "";
                    }
                }
                else
                {
                    T3RBoltStress_In_Label.Text = "";
                }
                T3RBoltStress_SI_Label.Text = SelectedApp_DataSet.Tables[0].Rows[0]["T3_ResidualBoltStress"].ToString().Trim();

                if ((SelectedApp_DataSet.Tables[0].Rows[0]["T3_ResidualBoltLoad"].ToString().Trim() != "") && (SelectedApp_DataSet.Tables[0].Rows[0]["T3_ResidualBoltLoad"].ToString().Trim().Length > 0))
                {
                    try
                    {
                        decimal T1_BPLTest = Convert.ToDecimal(SelectedApp_DataSet.Tables[0].Rows[0]["T3_ResidualBoltLoad"].ToString().Trim());
                        T3RBoltLoad_In_Label.Text = Math.Round(T1_BPLTest * 145.037738007M, 2, MidpointRounding.AwayFromZero).ToString();
                    }
                    catch (Exception T1_APStress_Exception)
                    {
                        MessageBox.Show(T1_APStress_Exception.Message + " while setting Bolt Stress Tab.");
                        T3RBoltLoad_In_Label.Text = "";
                    }
                }
                else
                {
                    T3RBoltLoad_In_Label.Text = "";
                }
                T3RBoltLoad_SI_Label.Text = SelectedApp_DataSet.Tables[0].Rows[0]["T3_ResidualBoltLoad"].ToString().Trim();
                T3RBoltYield_Per_Label.Text = SelectedApp_DataSet.Tables[0].Rows[0]["T3_ResidualBoltYieldPC"].ToString().Trim();

                // Row 5
                if ((SelectedApp_DataSet.Tables[0].Rows[0]["DetensioningBoltStress"].ToString().Trim() != "") && (SelectedApp_DataSet.Tables[0].Rows[0]["DetensioningBoltStress"].ToString().Trim().Length > 0))
                {
                    try
                    {
                        decimal T2_RTest = Convert.ToDecimal(SelectedApp_DataSet.Tables[0].Rows[0]["DetensioningBoltStress"].ToString().Trim());
                        DetenBoltStress_In_Label.Text = Math.Round(T2_RTest * 145.037738007M, 2, MidpointRounding.AwayFromZero).ToString();
                    }
                    catch (Exception T1_APStress_Exception)
                    {
                        MessageBox.Show(T1_APStress_Exception.Message + " while setting Bolt Stress Tab.");
                        DetenBoltStress_In_Label.Text = "";
                    }
                }
                else
                {
                    DetenBoltStress_In_Label.Text = "";
                }
                DetenBoltStress_SI_Label.Text = SelectedApp_DataSet.Tables[0].Rows[0]["DetensioningBoltStress"].ToString().Trim();

                if ((SelectedApp_DataSet.Tables[0].Rows[0]["DetensioningBoltLoad"].ToString().Trim() != "") && (SelectedApp_DataSet.Tables[0].Rows[0]["DetensioningBoltLoad"].ToString().Trim().Length > 0))
                {
                    try
                    {
                        decimal T1_BPLTest = Convert.ToDecimal(SelectedApp_DataSet.Tables[0].Rows[0]["DetensioningBoltLoad"].ToString().Trim());
                        DetenBoltLoad_In_Label.Text = Math.Round(T1_BPLTest * 145.037738007M, 2, MidpointRounding.AwayFromZero).ToString();
                    }
                    catch (Exception T1_APStress_Exception)
                    {
                        MessageBox.Show(T1_APStress_Exception.Message + " while setting Bolt Stress Tab.");
                        DetenBoltLoad_In_Label.Text = "";
                    }
                }
                else
                {
                    DetenBoltLoad_In_Label.Text = "";
                }
                DetenBoltLoad_SI_Label.Text = SelectedApp_DataSet.Tables[0].Rows[0]["DetensioningBoltLoad"].ToString().Trim();
                DetenBoltYield_Per_Label.Text = SelectedApp_DataSet.Tables[0].Rows[0]["DetensioningBoltYieldPC"].ToString().Trim();

                // Sequence Tab
                if ((SelectedApp_DataSet.Tables[0].Rows[0]["TensionPressure_FirstPass"].ToString().Trim() != "") && (SelectedApp_DataSet.Tables[0].Rows[0]["TensionPressure_FirstPass"].ToString().Trim().Length > 0))
                {
                    try
                    {
                        decimal Pass1Test = Convert.ToDecimal(SelectedApp_DataSet.Tables[0].Rows[0]["TensionPressure_FirstPass"].ToString());
                        Pass1AppliedPressure_SequenceTabValue_In_Label.Text = Math.Round((Convert.ToDecimal(SelectedApp_DataSet.Tables[0].Rows[0]["TensionPressure_FirstPass"].ToString()) * 145.037738007M), 0, MidpointRounding.AwayFromZero).ToString();
                    }
                    catch (Exception Pass1_Exception)
                    {
                        MessageBox.Show(Pass1_Exception.Message + "  First pass pressure found wrong while setting application.");
                        Pass1AppliedPressure_SequenceTabValue_In_Label.Text = "";
                    }
                }
                else
                {
                    Pass1AppliedPressure_SequenceTabValue_In_Label.Text = "";
                }
                Pass1AppliedPressure_SequenceTabValue_SI_Label.Text = SelectedApp_DataSet.Tables[0].Rows[0]["TensionPressure_FirstPass"].ToString();
                Pass2Bolt_SequenceTabValue_Label.Text = SelectedApp_DataSet.Tables[0].Rows[0]["Bolt02"].ToString();
                Pass2AppliedPressure_SequenceTabValue_SI_Label.Text = SelectedApp_DataSet.Tables[0].Rows[0]["TensionPressure_SecondPass"].ToString();
                if ((SelectedApp_DataSet.Tables[0].Rows[0]["TensionPressure_SecondPass"].ToString().Trim() != "") && (SelectedApp_DataSet.Tables[0].Rows[0]["TensionPressure_SecondPass"].ToString().Trim().Length > 0))
                {
                    try
                    {
                        decimal Pressure2_Test = Convert.ToDecimal(SelectedApp_DataSet.Tables[0].Rows[0]["TensionPressure_SecondPass"].ToString());
                        Pass2AppliedPressure_SequenceTabValue_In_Label.Text = Math.Round((Convert.ToDecimal(SelectedApp_DataSet.Tables[0].Rows[0]["TensionPressure_SecondPass"].ToString()) * 145.037738007M), 0, MidpointRounding.AwayFromZero).ToString();
                    }
                    catch (Exception Pressure2_Exception)
                    {
                        MessageBox.Show(Pressure2_Exception.Message + "  Second pass pressure found wrong while setting application.");
                        Pass2AppliedPressure_SequenceTabValue_In_Label.Text = "";
                    }
                }
                else
                {
                    Pass2AppliedPressure_SequenceTabValue_In_Label.Text = "";
                }
                Pass3Bolt_SequenceTabValue_Label.Text = SelectedApp_DataSet.Tables[0].Rows[0]["Bolt03"].ToString();
                Pass3AppliedPressure_SequenceTabValue_SI_Label.Text = SelectedApp_DataSet.Tables[0].Rows[0]["TensionPressure_ThirdPass"].ToString();
                Pass4Bolt_SequenceTabValue_Label.Text = SelectedApp_DataSet.Tables[0].Rows[0]["Bolt04"].ToString();
                Pass4AppliedPressure_SequenceTabValue_SI_Label.Text = SelectedApp_DataSet.Tables[0].Rows[0]["TensionPressure_ThirdPass"].ToString();
                if ((SelectedApp_DataSet.Tables[0].Rows[0]["TensionPressure_ThirdPass"].ToString().Trim() != "") && (SelectedApp_DataSet.Tables[0].Rows[0]["TensionPressure_ThirdPass"].ToString().Trim().Length > 0))
                {
                    try
                    {
                        decimal Pass3_Test = Convert.ToDecimal(SelectedApp_DataSet.Tables[0].Rows[0]["TensionPressure_ThirdPass"].ToString());
                        Pass3AppliedPressure_SequenceTabValue_In_Label.Text = Math.Round((Convert.ToDecimal(SelectedApp_DataSet.Tables[0].Rows[0]["TensionPressure_ThirdPass"].ToString()) * 145.037738007M), 0, MidpointRounding.AwayFromZero).ToString();
                        Pass4AppliedPressure_SequenceTabValue_In_Label.Text = Math.Round((Convert.ToDecimal(SelectedApp_DataSet.Tables[0].Rows[0]["TensionPressure_ThirdPass"].ToString()) * 145.037738007M), 0, MidpointRounding.AwayFromZero).ToString();
                    }
                    catch (Exception Pass3_Exception)
                    {
                        MessageBox.Show(Pass3_Exception.Message + " Pass 3 pressure found wrong while setting application.");
                        Pass3AppliedPressure_SequenceTabValue_In_Label.Text = "";
                        Pass4AppliedPressure_SequenceTabValue_In_Label.Text = "";
                    }
                }
                else
                {
                    Pass3AppliedPressure_SequenceTabValue_In_Label.Text = "";
                    Pass4AppliedPressure_SequenceTabValue_In_Label.Text = "";
                }
                CheckPass1AppliedPressure_SequenceTabValue_SI_Label.Text = SelectedApp_DataSet.Tables[0].Rows[0]["CheckingPass"].ToString();
                CheckPass2AppliedPressure_SequenceTabValue_SI_Label.Text = SelectedApp_DataSet.Tables[0].Rows[0]["CheckingPass"].ToString();
                if ((SelectedApp_DataSet.Tables[0].Rows[0]["CheckingPass"].ToString().Trim() != "") && (SelectedApp_DataSet.Tables[0].Rows[0]["CheckingPass"].ToString().Trim().Length > 0))
                {
                    try
                    {
                        decimal Check_Test = Convert.ToDecimal(SelectedApp_DataSet.Tables[0].Rows[0]["CheckingPass"].ToString());
                        CheckPass1AppliedPressure_SequenceTabValue_In_Label.Text = Math.Round((Convert.ToDecimal(SelectedApp_DataSet.Tables[0].Rows[0]["CheckingPass"].ToString()) * 145.037738007M), 0, MidpointRounding.AwayFromZero).ToString();
                        CheckPass2AppliedPressure_SequenceTabValue_In_Label.Text = Math.Round((Convert.ToDecimal(SelectedApp_DataSet.Tables[0].Rows[0]["CheckingPass"].ToString()) * 145.037738007M), 0, MidpointRounding.AwayFromZero).ToString();
                    }
                    catch (Exception Check_Exception)
                    {
                        MessageBox.Show(Check_Exception.Message + " Checking pressure found wrong while setting application.");
                        CheckPass1AppliedPressure_SequenceTabValue_In_Label.Text = "";
                        CheckPass2AppliedPressure_SequenceTabValue_In_Label.Text = "";
                    }
                }
                else
                {
                    CheckPass1AppliedPressure_SequenceTabValue_In_Label.Text = "";
                    CheckPass2AppliedPressure_SequenceTabValue_In_Label.Text = "";
                }
                // Bolt Tab
                BoltMaterialValueTab_Label.Text = SelectedApp_DataSet.Tables[0].Rows[0]["Material"].ToString().Trim();
                NominalBoltDiameterValueTab_Label.Text = SelectedApp_DataSet.Tables[0].Rows[0]["BoltThread"].ToString().Trim();
                NumberOfBoltsValueTab_Label.Text = SelectedApp_DataSet.Tables[0].Rows[0]["Number_of_Bolts"].ToString().Trim();
                
                // Info Panel
                FlangeName_Label.Text = SelectedApp_DataSet.Tables[0].Rows[0]["FlangeRating"].ToString().Trim();
                Id_Label.Text = SelectedApp_DataSet.Tables[0].Rows[0]["JointId"].ToString().Trim();
                SpecificationInfo_Label.Text = SelectedApp_DataSet.Tables[0].Rows[0]["Specification"].ToString().Trim();
                string InfoPanel_BoltThread = SelectedApp_DataSet.Tables[0].Rows[0]["BoltThread"].ToString().Trim();
                if (InfoPanel_BoltThread.Substring(0, 1) == "M")
                {
                    NominalThreadSizeInfo_Label.Text = InfoPanel_BoltThread + "x" + SelectedApp_DataSet.Tables[0].Rows[0]["Pitch_TPI_Value"].ToString().Trim();
                    TPI_Label.Text = "Pitch";
                }
                else
                {
                    NominalThreadSizeInfo_Label.Text = InfoPanel_BoltThread + "-" + SelectedApp_DataSet.Tables[0].Rows[0]["Pitch_TPI_Value"].ToString().Trim() +"UN";
                    TPI_Label.Text = "TPI";
                }
                PitchInfo_Label.Text = SelectedApp_DataSet.Tables[0].Rows[0]["Pitch_TPI_Value"].ToString().Trim();
                NoOfBoltsInfo_Label.Text = SelectedApp_DataSet.Tables[0].Rows[0]["Number_of_Bolts"].ToString().Trim();
                BoltToolInfo_Label.Text = SelectedApp_DataSet.Tables[0].Rows[0]["Bolt_to_ToolRatio"].ToString().Trim();
                if (SelectedApp_DataSet.Tables[0].Rows[0]["TotalClampLength"].ToString().Trim().Length > 0)
                {
                    try
                    {
                        InfoPanel_ClampLength = Convert.ToDecimal(SelectedApp_DataSet.Tables[0].Rows[0]["TotalClampLength"].ToString().Trim());
                    }
                    catch (Exception InfoPanel_ClampLength_Exception)
                    {
                        MessageBox.Show(InfoPanel_ClampLength_Exception.Message + " while setting Info Panel.");
                    }
                }
                ClampLengthInfo_Label.Text = Math.Round((InfoPanel_ClampLength / 25.4M), 4, MidpointRounding.AwayFromZero).ToString() + " in";
                ClampLengthInfo_mm_Label.Text = SelectedApp_DataSet.Tables[0].Rows[0]["TotalClampLength"].ToString().Trim() + " mm";
                LTF_Label.Text = SelectedApp_DataSet.Tables[0].Rows[0]["LTF"].ToString().Trim();
                ToolId_Info_Label.Text = SpecificationInfo_Label.Text = SelectedApp_DataSet.Tables[0].Rows[0]["ModelNumber"].ToString().Trim();
                TensionerBolt_Info_Label.Text = NominalThreadSizeInfo_Label.Text;
                if (SelectedApp_DataSet.Tables[0].Rows[0]["ToolPressureArea"].ToString().Trim().Length > 0)
                {
                    try
                    {
                        InfoPanel_ToolPressureArea = Convert.ToDecimal(SelectedApp_DataSet.Tables[0].Rows[0]["ToolPressureArea"].ToString().Trim());
                    }
                    catch (Exception InfoPanel_ToolPressureArea_Exception)
                    {
                        MessageBox.Show(InfoPanel_ToolPressureArea_Exception.Message + " while setting Info Panel.");
                    }
                }
                ToolPressureArea_Info_Label.Text = Math.Round((InfoPanel_ToolPressureArea / 645.16M), 4, MidpointRounding.AwayFromZero).ToString() + " Sq in";
                ToolPressureArea_mm_Label.Text = SelectedApp_DataSet.Tables[0].Rows[0]["ToolPressureArea"].ToString().Trim() + " Sq mm";
                BoltMaterial_Info_Label.Text = SelectedApp_DataSet.Tables[0].Rows[0]["Material"].ToString().Trim();
                if (SelectedApp_DataSet.Tables[0].Rows[0]["BoltYield"].ToString().Trim().Length > 0)
                {
                    try
                    {
                        InfoPanel_BoltYield = Convert.ToDecimal(SelectedApp_DataSet.Tables[0].Rows[0]["BoltYield"].ToString().Trim());
                    }
                    catch (Exception InfoPanel_BoltYield_Exception)
                    {
                        MessageBox.Show(InfoPanel_BoltYield_Exception.Message + " while setting Info Panel.");
                    }
                }
                BoltYield_In_Label.Text = Math.Round((InfoPanel_BoltYield * 0.145037738007M), 4, MidpointRounding.AwayFromZero).ToString().Trim() + " ksi";
                BoltYieldStrengthValueTab_In_Label.Text = BoltYield_In_Label.Text;          // Bolt Tab
                BoltYield_SI_Label.Text = SelectedApp_DataSet.Tables[0].Rows[0]["BoltYield"].ToString().Trim() + " MPa";
                BoltYieldStrengthValueTab_SI_Label.Text = BoltYield_SI_Label.Text;          // Bolt Tab

                if (SelectedApp_DataSet.Tables[0].Rows[0]["BoltTensileStressArea"].ToString().Trim().Length > 0)
                {
                    try
                    {
                        InfoPanel_TensileStressArea = Convert.ToDecimal(SelectedApp_DataSet.Tables[0].Rows[0]["BoltTensileStressArea"].ToString().Trim());
                    }
                    catch (Exception InfoPanel_TensileStressArea_Exception)
                    {
                        MessageBox.Show(InfoPanel_TensileStressArea_Exception.Message + " while setting Info Panel.");
                    }
                }
                TensileStressArea_In_Label.Text = Math.Round((InfoPanel_TensileStressArea / 645.16M), 4, MidpointRounding.AwayFromZero).ToString().Trim() + " Sq in";
                TensileStressArea_SI_Label.Text = SelectedApp_DataSet.Tables[0].Rows[0]["BoltTensileStressArea"].ToString().Trim() + " Sq mm";
                TensileStressAreaValueTab_In_Label.Text = TensileStressArea_In_Label.Text;      // Bolt Tab
                TensileStressAreaValueTab_SI_Label.Text = TensileStressArea_SI_Label.Text;      // Bolt Tab
                
                if (SelectedApp_DataSet.Tables[0].Rows[0]["BoltLength"].ToString().Trim().Length > 0)
                {
                    try
                    {
                        InfoPanel_BoltLength = Convert.ToDecimal(SelectedApp_DataSet.Tables[0].Rows[0]["BoltLength"].ToString().Trim());
                    }
                    catch (Exception InfoPanel_BoltLength_Exception)
                    {
                        MessageBox.Show(InfoPanel_BoltLength_Exception.Message + " while setting Info Panel.");
                    }
                }
                BoltLength_In_Label.Text = Math.Round((InfoPanel_BoltLength / 25.4M), 4, MidpointRounding.AwayFromZero).ToString() + " in";
                BoltLength_SI_Label.Text = SelectedApp_DataSet.Tables[0].Rows[0]["BoltLength"].ToString().Trim() + " mm";
                BoltLengthValueTab_In_Label.Text = BoltLength_In_Label.Text;            // Bolt Tab
                BoltLengthValueTab_SI_Label.Text = BoltLength_SI_Label.Text;            // Bolt Tab
                // Pressure A & B is handled above in Pressure.
                // Torque value is also set above

                //if (SelectedApp_DataSet.Tables[0].Rows[0]["Torque"].ToString().Trim().Length > 0)
                //{
                //    try
                //    {
                //        InfoPanel_Torque = Convert.ToDecimal(SelectedApp_DataSet.Tables[0].Rows[0]["Torque"].ToString().Trim());
                //    }
                //    catch (Exception InfoPanel_Torque_Exception)
                //    {
                //        MessageBox.Show(InfoPanel_Torque_Exception.Message + " while setting Info Panel.");
                //    }
                //}
                //Torque_In_Label.Text = Math.Round(InfoPanel_Torque / 1.35581794833M, 0, MidpointRounding.AwayFromZero).ToString() + " lb-ft";
                //Torque_SI_Label.Text = SelectedApp_DataSet.Tables[0].Rows[0]["Torque"].ToString().Trim() + " N-m";
                CoefficientValue_Label.Text = SelectedApp_DataSet.Tables[0].Rows[0]["Coefficient_Friction"].ToString().Trim();

                // Sequence Tab
                int Bolt_Tool_Ratio = Convert.ToInt32(SelectedApp_DataSet.Tables[0].Rows[0]["Bolt_to_ToolRatioId"]);
                switch (Bolt_Tool_Ratio)
                {
                    case 1:
                        int Bolts100 = Convert.ToInt32(NumberOfBolts_TextBox.Text);
                        switch (Bolts100)
                        {
                            case 4:
                                Sequence_PictureBox.Image = WizBolt.Properties.Resources.Flange100_04;
                                break;
                            case 8:
                                Sequence_PictureBox.Image = WizBolt.Properties.Resources.Flange100_08;
                                break;
                            case 12:
                                Sequence_PictureBox.Image = WizBolt.Properties.Resources.Flange100_12;
                                break;
                            case 16:
                                Sequence_PictureBox.Image = WizBolt.Properties.Resources.Flange100_16;
                                break;
                            case 20:
                                Sequence_PictureBox.Image = WizBolt.Properties.Resources.Flange100_20;
                                break;
                            case 24:
                                Sequence_PictureBox.Image = WizBolt.Properties.Resources.Flange100_24;
                                break;
                            case 28:
                                Sequence_PictureBox.Image = WizBolt.Properties.Resources.Flange100_28;
                                break;
                            case 32:
                                Sequence_PictureBox.Image = WizBolt.Properties.Resources.Flange100_32;
                                break;
                            case 36:
                                Sequence_PictureBox.Image = WizBolt.Properties.Resources.Flange100_36;
                                break;
                            case 40:
                                Sequence_PictureBox.Image = WizBolt.Properties.Resources.Flange100_40;
                                break;
                            case 44:
                                Sequence_PictureBox.Image = WizBolt.Properties.Resources.Flange100_44;
                                break;
                            case 48:
                                Sequence_PictureBox.Image = WizBolt.Properties.Resources.Flange100_48;
                                break;
                            case 52:
                                Sequence_PictureBox.Image = WizBolt.Properties.Resources.Flange100_52;
                                break;
                            case 56:
                                Sequence_PictureBox.Image = WizBolt.Properties.Resources.Flange100_56;
                                break;
                            case 60:
                                Sequence_PictureBox.Image = WizBolt.Properties.Resources.Flange100_60;
                                break;
                            default:
                                Sequence_PictureBox.Image = WizBolt.Properties.Resources.Sequence_Full;
                                break;
                        }
                        break;
                    case 2:
                        int TorquingBoltImages = Convert.ToInt32(NumberOfBolts_TextBox.Text);
                        switch(TorquingBoltImages)
                        {
                            case 4:
                                Sequence_PictureBox.Image = WizBolt.Properties.Resources.Flange050_04;
                                break;
                            case 8:
                                Sequence_PictureBox.Image = WizBolt.Properties.Resources.Flange050_08;
                                break;
                            case 12:
                                Sequence_PictureBox.Image = WizBolt.Properties.Resources.Flange050_12;
                                break;
                            case 16:
                                Sequence_PictureBox.Image = WizBolt.Properties.Resources.Flange050_16;
                                break;
                            case 20:
                                Sequence_PictureBox.Image = WizBolt.Properties.Resources.Flange050_20;
                                break;
                            case 24:
                                Sequence_PictureBox.Image = WizBolt.Properties.Resources.Flange050_24;
                                break;
                            case 28:
                                Sequence_PictureBox.Image = WizBolt.Properties.Resources.Flange050_28;
                                break;
                            case 32:
                                Sequence_PictureBox.Image = WizBolt.Properties.Resources.Flange050_32;
                                break;
                            case 36:
                                Sequence_PictureBox.Image = WizBolt.Properties.Resources.Flange050_36;
                                break;
                            case 40:
                                Sequence_PictureBox.Image = WizBolt.Properties.Resources.Flange050_40;
                                break;
                            case 44:
                                Sequence_PictureBox.Image = WizBolt.Properties.Resources.Flange050_44;
                                break;
                            case 48:
                                Sequence_PictureBox.Image = WizBolt.Properties.Resources.Flange050_48;
                                break;
                            case 52:
                                Sequence_PictureBox.Image = WizBolt.Properties.Resources.Flange050_52;
                                break;
                            case 56:
                                Sequence_PictureBox.Image = WizBolt.Properties.Resources.Flange050_56;
                                break;
                            case 60:
                                Sequence_PictureBox.Image = WizBolt.Properties.Resources.Flange050_60;
                                break;
                            default:
                                Sequence_PictureBox.Image = WizBolt.Properties.Resources.Sequence_half;
                                break;
                        }
                        break;
                    case 3:
                        int Bolts25 = Convert.ToInt32(NumberOfBolts_TextBox.Text);
                        switch (Bolts25)
                        {
                            case 4:
                                Sequence_PictureBox.Image = WizBolt.Properties.Resources.Flange025_04;
                                break;
                            case 8:
                                Sequence_PictureBox.Image = WizBolt.Properties.Resources.Flange025_08;
                                break;
                            case 12:
                                Sequence_PictureBox.Image = WizBolt.Properties.Resources.Flange025_12;
                                break;
                            case 16:
                                Sequence_PictureBox.Image = WizBolt.Properties.Resources.Flange025_16;
                                break;
                            case 20:
                                Sequence_PictureBox.Image = WizBolt.Properties.Resources.Flange025_20;
                                break;
                            case 24:
                                Sequence_PictureBox.Image = WizBolt.Properties.Resources.Flange025_24;
                                break;
                            case 28:
                                Sequence_PictureBox.Image = WizBolt.Properties.Resources.Flange025_28;
                                break;
                            case 32:
                                Sequence_PictureBox.Image = WizBolt.Properties.Resources.Flange025_32;
                                break;
                            case 36:
                                Sequence_PictureBox.Image = WizBolt.Properties.Resources.Flange025_36;
                                break;
                            case 40:
                                Sequence_PictureBox.Image = WizBolt.Properties.Resources.Flange025_40;
                                break;
                            case 44:
                                Sequence_PictureBox.Image = WizBolt.Properties.Resources.Flange025_44;
                                break;
                            case 48:
                                Sequence_PictureBox.Image = WizBolt.Properties.Resources.Flange025_48;
                                break;
                            case 52:
                                Sequence_PictureBox.Image = WizBolt.Properties.Resources.Flange025_52;
                                break;
                            case 56:
                                Sequence_PictureBox.Image = WizBolt.Properties.Resources.Flange025_56;
                                break;
                            case 60:
                                Sequence_PictureBox.Image = WizBolt.Properties.Resources.Flange025_60;
                                break;
                            default:
                                Sequence_PictureBox.Image = WizBolt.Properties.Resources.Sequence_quarter;
                                break;
                        }
                        break;

                    case 4:
                        int TorqueBolts = Convert.ToInt32(NumberOfBolts_TextBox.Text);
                        switch (TorqueBolts)
                        {
                            case 4:
                                Sequence_PictureBox.Image = WizBolt.Properties.Resources.FlangeTorque_04;
                                break;
                            case 8:
                                Sequence_PictureBox.Image = WizBolt.Properties.Resources.FlangeTorque_08;
                                break;
                            case 12:
                                Sequence_PictureBox.Image = WizBolt.Properties.Resources.FlangeTorque_12;
                                break;
                            case 16:
                                Sequence_PictureBox.Image = WizBolt.Properties.Resources.FlangeTorque_16;
                                break;
                            case 20:
                                Sequence_PictureBox.Image = WizBolt.Properties.Resources.FlangeTorque_20;
                                break;
                            case 24:
                                Sequence_PictureBox.Image = WizBolt.Properties.Resources.FlangeTorque_24;
                                break;
                            case 28:
                                Sequence_PictureBox.Image = WizBolt.Properties.Resources.FlangeTorque_28;
                                break;
                            case 32:
                                Sequence_PictureBox.Image = WizBolt.Properties.Resources.FlangeTorque_32;
                                break;
                            case 36:
                                Sequence_PictureBox.Image = WizBolt.Properties.Resources.FlangeTorque_36;
                                break;
                            case 40:
                                Sequence_PictureBox.Image = WizBolt.Properties.Resources.FlangeTorque_40;
                                break;
                            case 44:
                                Sequence_PictureBox.Image = WizBolt.Properties.Resources.FlangeTorque_44;
                                break;
                            case 48:
                                Sequence_PictureBox.Image = WizBolt.Properties.Resources.FlangeTorque_48;
                                break;
                            case 52:
                                Sequence_PictureBox.Image = WizBolt.Properties.Resources.FlangeTorque_52;
                                break;
                            case 56:
                                Sequence_PictureBox.Image = WizBolt.Properties.Resources.FlangeTorque_56;
                                break;
                            case 60:
                                Sequence_PictureBox.Image = WizBolt.Properties.Resources.FlangeTorque_60;
                                break;
                            default:
                                Sequence_PictureBox.Image = null;
                                break;
                        }
                        break;
                }
                // Graph Tab
                DrawGraph();
            }
            else
            {
                MessageBox.Show("This application may not be available or not found. Was it saved properly?", "Error!");
            }

            IsSetting = false;
        }

        public void SaveProject()
        {
            // This is bit strange, but first file is created, obviously empty that means file name is written on hard disk.
            SaveFileDialog SaveProjectFile = new SaveFileDialog();
            SaveProjectFile.Filter = "Project Files (*.wizbolt)|*.wizbolt";
            SaveProjectFile.DefaultExt = "wizbolt";
            SaveProjectFile.AddExtension = true;
            SaveProjectFile.RestoreDirectory = true;
            if (FileWithPath_Label.Text == "File with path")
            {
                DialogResult SaveResult = SaveProjectFile.ShowDialog();
                if (SaveResult == DialogResult.OK)
                {
                    ProjectSavedFile = SaveProjectFile.FileName;
                    FileWithPath_Label.Text = ProjectSavedFile;
                    Write_DiskFile();
                }
            }
            else
            {
                Write_DiskFile();
            }
        }

        public void SaveAs()
        {
            SaveFileDialog SaveProjectFile = new SaveFileDialog();
            SaveProjectFile.Filter = "Project Files (*.wizbolt)|*.wizbolt";
            SaveProjectFile.DefaultExt = "wizbolt";
            SaveProjectFile.AddExtension = true;
            SaveProjectFile.RestoreDirectory = true;
            DialogResult SaveResult = SaveProjectFile.ShowDialog();
            if (SaveResult == DialogResult.OK)
            {
                ProjectSavedFile = SaveProjectFile.FileName;
                FileWithPath_Label.Text = ProjectSavedFile;
                Write_DiskFile();
            }

        }
        protected void Write_DiskFile()
        {
            string[] ProjectStartDate = new string[3] {"00", "00", "00"};
            string[] ProjectEndDate = new string[3] { "00", "00", "00" };
            if ((FileWithPath_Label.Text.Trim() != "File with path") && (FileWithPath_Label.Text.Trim().Length > 0))
            {
                ProjectSavedFile = FileWithPath_Label.Text.Trim();
            }
            else
                if ((ProjectSavedFile == string.Empty) || (ProjectSavedFile.Length == 0))
                {
                    MessageBox.Show("File or file path not found. Cannot save project.", "Error!");
                }
                else
                {
                    FileWithPath_Label.Text = ProjectSavedFile;
                }
            if (EnteredProjectDate_label.Text.Trim().Contains('-'))
            {
                ProjectStartDate = EnteredProjectDate_label.Text.Trim().Split('-');
            }
            else
            {
                ProjectStartDate = EnteredProjectDate_label.Text.Trim().Split('/');
            }
            if (ProjectEndDate_Label.Text.Trim() != "Not Entered")
            {
                if (ProjectEndDate_Label.Text.Trim().Contains('-'))
                {
                    ProjectEndDate = ProjectEndDate_Label.Text.Trim().Split('-');
                }
                else
                {
                    ProjectEndDate = ProjectEndDate_Label.Text.Trim().Split('/');
                }
            }

            FileStream Project_FileStream = new FileStream(ProjectSavedFile, FileMode.Create);
            StreamWriter Project_Writer = new StreamWriter(Project_FileStream);
            try
            {
                //Project_Writer.Write(ProjectId_Label.Text);
                //Project_Writer.Write("~");
                //Project_Writer.Write(CustomerId_Label.Text);
                        

                // Project//
                Project_Writer.WriteLine(Encrypt("!!~!!Wizodyssey//~//", "PolarissimaDwarf", "WizERPEmance", "SHA1", 10));   // Title 
                Project_Writer.WriteLine(Encrypt("!!~!!Project_Start//~//", "PolarissimaDwarf", "WizERPEmance", "SHA1", 10));
                Project_Writer.WriteLine(Encrypt("!!~!!0001%^%" + CustomerName_Label.Text.Trim() + "//~//", "PolarissimaDwarf", "WizERPEmance", "SHA1", 10));
                Project_Writer.WriteLine(Encrypt("!!~!!0002%^%" + LocationName_Label.Text.Trim() + "//~//", "PolarissimaDwarf", "WizERPEmance", "SHA1", 10));
                Project_Writer.WriteLine(Encrypt("!!~!!0003%^%" + ProjectName_Label.Text.Trim() + "//~//", "PolarissimaDwarf", "WizERPEmance", "SHA1", 10));
                Project_Writer.WriteLine(Encrypt("!!~!!0004%^%" + Reference_Label.Text.Trim() + "//~//", "PolarissimaDwarf", "WizERPEmance", "SHA1", 10));
                Project_Writer.WriteLine(Encrypt("!!~!!0005%^%" + ProjectStartDate[0].Trim() + "//~//", "PolarissimaDwarf", "WizERPEmance", "SHA1", 10));
                Project_Writer.WriteLine(Encrypt("!!~!!0006%^%" + ProjectStartDate[1].Trim() + "//~//", "PolarissimaDwarf", "WizERPEmance", "SHA1", 10));
                Project_Writer.WriteLine(Encrypt("!!~!!0007%^%" + ProjectStartDate[2].Trim() + "//~//", "PolarissimaDwarf", "WizERPEmance", "SHA1", 10));
                Project_Writer.WriteLine(Encrypt("!!~!!0008%^%" + ProjectEndDate[0].Trim() + "//~//", "PolarissimaDwarf", "WizERPEmance", "SHA1", 10));
                Project_Writer.WriteLine(Encrypt("!!~!!0009%^%" + ProjectEndDate[1].Trim() + "//~//", "PolarissimaDwarf", "WizERPEmance", "SHA1", 10));
                Project_Writer.WriteLine(Encrypt("!!~!!0010%^%" + ProjectEndDate[2].Trim() + "//~//", "PolarissimaDwarf", "WizERPEmance", "SHA1", 10));
                Project_Writer.WriteLine(Encrypt("!!~!!0011%^%" + Engineer_Label.Text.Trim() + "//~//", "PolarissimaDwarf", "WizERPEmance", "SHA1", 10));
                Project_Writer.WriteLine(Encrypt("!!~!!0012%^%" + Notes_RichTextBox.Text + "//~//", "PolarissimaDwarf", "WizERPEmance", "SHA1", 10));
                Project_Writer.WriteLine(Encrypt("!!~!!0013%^%" + Summary_Label.Text.Trim() + "//~//", "PolarissimaDwarf", "WizERPEmance", "SHA1", 10));
                Project_Writer.WriteLine(Encrypt("!!~!!0014%^%" + ToolSeriesId_Label.Text.Trim() + "//~//", "PolarissimaDwarf", "WizERPEmance", "SHA1", 10));
                Project_Writer.WriteLine(Encrypt("!!~!!0015%^%" + ToolSeries_Label.Text.Trim() + "//~//", "PolarissimaDwarf", "WizERPEmance", "SHA1", 10));
                Project_Writer.WriteLine(Encrypt("!!~!!0016%^%" + CrossLoading_TextBox.Text.Trim() + "//~//", "PolarissimaDwarf", "WizERPEmance", "SHA1", 10));
                Project_Writer.WriteLine(Encrypt("!!~!!0017%^%" + Detensioning_TextBox.Text.Trim() + "//~//", "PolarissimaDwarf", "WizERPEmance", "SHA1", 10));
                Project_Writer.WriteLine(Encrypt("!!~!!0018%^%" + CoefficientValue_Label.Text.Trim() + "//~//", "PolarissimaDwarf", "WizERPEmance", "SHA1", 10));
                Project_Writer.WriteLine(Encrypt("!!~!!0019%^%" + Stress_Label.Text.Trim() + "//~//", "PolarissimaDwarf", "WizERPEmance", "SHA1", 10));
                Project_Writer.WriteLine(Encrypt("!!~!!Project_End//~//", "PolarissimaDwarf", "WizERPEmance", "SHA1", 10));
            }
            catch (Exception ProjectWrite_Exception)
            {
                MessageBox.Show(ProjectWrite_Exception.Message + " Error occurred while writing project file.", "Error!");
            }

            SQLiteConnection SaveProject_Connection = new SQLiteConnection("Data Source=WizBolt.db;New=False;");
            SQLiteCommand SaveProject_Command = new SQLiteCommand(SaveProject_Connection);

            SaveProject_Command.CommandText = " SELECT ProjectReportId, ApplicationId, UnitSystem, JointId, PrimaryStandardId,  "
                                            + "   Specification, PrimaryFlangeRatingId, FlangeRating, Flange1_TypeId,  "
                                            + "   Flange1_Abbreviation, Flange1_ClampLength, Flange2_TypeId, Flange2_Abbreviation,  "
                                            + "   Flange2_ClampLength, GasketId, Gasket, GasketGap,  "
                                            + "   SpacerId, Spacer, SpacerThickness, TotalClampLength, PrimaryFlangeBoltId,  "
                                            + "   BoltThread, Pitch_or_TPI, Pitch_TPI_Value, Number_of_Bolts, LTF, "
                                            + "   Bolt_to_ToolRatioId, Bolt_to_ToolRatio, ToolId, ModelNumber,  "
                                            + "   ToolPressureArea, MaterialId, Material, BoltYield, BoltTensileStressArea, "
                                            + "   BoltLength, BoltStressBase, BoltMinorDiameterArea, T1_APressureBoltStress, T1_APressureBoltLoad,  "
                                            + "   T1_APressureBoltYieldPC, T1_BPressureBoltStress, T1_BPressureBoltLoad, T1_BPressureBoltYieldPC,   "
                                            + "   T2_ResidualBoltStress, T2_ResidualBoltLoad, T2_ResidualBoltYieldPC, DetensioningBoltStress, DetensioningBoltLoad,	"
                                            + "   DetensioningBoltYieldPC, TensionPressure_FirstPass, TensionPressure_SecondPass, TensionPressure_ThirdPass, CheckingPass,   "
                                            + "   Torque, Coefficient_Friction, Bolt01, Bolt02, Bolt03, Bolt04, "
                                            + "   FirstPass_100, FirstPass_50, SecondPass_50, Max_Detensioning, "
                                            + "   ResidualBoltStress, Comments, IsDetension, CrossLoadingPC, DetensioningPC, T3_ResidualBoltStress, T3_ResidualBoltLoad,   "
                                            + "   T3_ResidualBoltYieldPC "
                                            + "   FROM ProjectDetailedReport "
                                            + "           WHERE        ProjectReportId =  " + ProjectId_Label.Text.Trim();
                                                   
            DataSet SaveProject_DataSet = new DataSet();
            SQLiteDataAdapter SaveProject_DataAdapter = new SQLiteDataAdapter(SaveProject_Command);

            try
            {
                int SaveProjectResult = SaveProject_DataAdapter.Fill(SaveProject_DataSet);
                if (SaveProjectResult > 0)
                {
                    Project_Writer.WriteLine(Encrypt("!!~!!Application_Start//~//", "PolarissimaDwarf", "WizERPEmance", "SHA1", 10));
                    int AppRec = SaveProject_DataSet.Tables[0].Rows.Count;
                    Project_Writer.WriteLine(Encrypt("!!~!!" + AppRec.ToString() + "//~//", "PolarissimaDwarf", "WizERPEmance", "SHA1", 10));
                    for (int AppRec_Counter = 0; AppRec_Counter < AppRec; AppRec_Counter++)
                    {
                        Project_Writer.WriteLine(Encrypt("!!~!!AppStart " + (AppRec_Counter + 1).ToString() + "//~//", "PolarissimaDwarf", "WizERPEmance", "SHA1", 10));
                        Project_Writer.WriteLine(Encrypt("!!~!!00001%^%" + SaveProject_DataSet.Tables[0].Rows[AppRec_Counter]["ApplicationId"].ToString().Trim() + "//~//", "PolarissimaDwarf", "WizERPEmance", "SHA1", 10));
                        Project_Writer.WriteLine(Encrypt("!!~!!00002%^%" + SaveProject_DataSet.Tables[0].Rows[AppRec_Counter]["UnitSystem"].ToString().Trim() + "//~//", "PolarissimaDwarf", "WizERPEmance", "SHA1", 10));
                        Project_Writer.WriteLine(Encrypt("!!~!!00003%^%" + SaveProject_DataSet.Tables[0].Rows[AppRec_Counter]["JointId"].ToString().Trim() + "//~//", "PolarissimaDwarf", "WizERPEmance", "SHA1", 10));
                        Project_Writer.WriteLine(Encrypt("!!~!!00004%^%" + SaveProject_DataSet.Tables[0].Rows[AppRec_Counter]["PrimaryStandardId"].ToString().Trim() + "//~//", "PolarissimaDwarf", "WizERPEmance", "SHA1", 10));
                        Project_Writer.WriteLine(Encrypt("!!~!!00005%^%" + SaveProject_DataSet.Tables[0].Rows[AppRec_Counter]["Specification"].ToString().Trim() + "//~//", "PolarissimaDwarf", "WizERPEmance", "SHA1", 10));
                        Project_Writer.WriteLine(Encrypt("!!~!!00006%^%" + SaveProject_DataSet.Tables[0].Rows[AppRec_Counter]["PrimaryFlangeRatingId"].ToString().Trim() + "//~//", "PolarissimaDwarf", "WizERPEmance", "SHA1", 10));
                        Project_Writer.WriteLine(Encrypt("!!~!!00007%^%" + SaveProject_DataSet.Tables[0].Rows[AppRec_Counter]["FlangeRating"].ToString().Trim() + "//~//", "PolarissimaDwarf", "WizERPEmance", "SHA1", 10));
                        Project_Writer.WriteLine(Encrypt("!!~!!00008%^%" + SaveProject_DataSet.Tables[0].Rows[AppRec_Counter]["Flange1_TypeId"].ToString().Trim() + "//~//", "PolarissimaDwarf", "WizERPEmance", "SHA1", 10));
                        Project_Writer.WriteLine(Encrypt("!!~!!00009%^%" + SaveProject_DataSet.Tables[0].Rows[AppRec_Counter]["Flange1_Abbreviation"].ToString().Trim() + "//~//", "PolarissimaDwarf", "WizERPEmance", "SHA1", 10));
                        Project_Writer.WriteLine(Encrypt("!!~!!00010%^%" + SaveProject_DataSet.Tables[0].Rows[AppRec_Counter]["Flange1_ClampLength"].ToString().Trim() + "//~//", "PolarissimaDwarf", "WizERPEmance", "SHA1", 10));
                        Project_Writer.WriteLine(Encrypt("!!~!!00011%^%" + SaveProject_DataSet.Tables[0].Rows[AppRec_Counter]["Flange2_TypeId"].ToString().Trim() + "//~//", "PolarissimaDwarf", "WizERPEmance", "SHA1", 10));
                        Project_Writer.WriteLine(Encrypt("!!~!!00012%^%" + SaveProject_DataSet.Tables[0].Rows[AppRec_Counter]["Flange2_Abbreviation"].ToString().Trim() + "//~//", "PolarissimaDwarf", "WizERPEmance", "SHA1", 10));
                        Project_Writer.WriteLine(Encrypt("!!~!!00013%^%" + SaveProject_DataSet.Tables[0].Rows[AppRec_Counter]["Flange2_ClampLength"].ToString().Trim() + "//~//", "PolarissimaDwarf", "WizERPEmance", "SHA1", 10));
                        Project_Writer.WriteLine(Encrypt("!!~!!00014%^%" + SaveProject_DataSet.Tables[0].Rows[AppRec_Counter]["GasketId"].ToString().Trim() + "//~//", "PolarissimaDwarf", "WizERPEmance", "SHA1", 10));
                        Project_Writer.WriteLine(Encrypt("!!~!!00015%^%" + SaveProject_DataSet.Tables[0].Rows[AppRec_Counter]["Gasket"].ToString().Trim() + "//~//", "PolarissimaDwarf", "WizERPEmance", "SHA1", 10));
                        Project_Writer.WriteLine(Encrypt("!!~!!00016%^%" + SaveProject_DataSet.Tables[0].Rows[AppRec_Counter]["GasketGap"].ToString().Trim() + "//~//", "PolarissimaDwarf", "WizERPEmance", "SHA1", 10));
                        Project_Writer.WriteLine(Encrypt("!!~!!00017%^%" + SaveProject_DataSet.Tables[0].Rows[AppRec_Counter]["SpacerId"].ToString().Trim() + "//~//", "PolarissimaDwarf", "WizERPEmance", "SHA1", 10));
                        Project_Writer.WriteLine(Encrypt("!!~!!00018%^%" + SaveProject_DataSet.Tables[0].Rows[AppRec_Counter]["Spacer"].ToString().Trim() + "//~//", "PolarissimaDwarf", "WizERPEmance", "SHA1", 10));
                        Project_Writer.WriteLine(Encrypt("!!~!!00019%^%" + SaveProject_DataSet.Tables[0].Rows[AppRec_Counter]["SpacerThickness"].ToString().Trim() + "//~//", "PolarissimaDwarf", "WizERPEmance", "SHA1", 10));
                        Project_Writer.WriteLine(Encrypt("!!~!!00020%^%" + SaveProject_DataSet.Tables[0].Rows[AppRec_Counter]["TotalClampLength"].ToString().Trim() + "//~//", "PolarissimaDwarf", "WizERPEmance", "SHA1", 10));
                        Project_Writer.WriteLine(Encrypt("!!~!!00021%^%" + SaveProject_DataSet.Tables[0].Rows[AppRec_Counter]["PrimaryFlangeBoltId"].ToString().Trim() + "//~//", "PolarissimaDwarf", "WizERPEmance", "SHA1", 10));
                        Project_Writer.WriteLine(Encrypt("!!~!!00022%^%" + SaveProject_DataSet.Tables[0].Rows[AppRec_Counter]["BoltThread"].ToString().Trim() + "//~//", "PolarissimaDwarf", "WizERPEmance", "SHA1", 10));
                        Project_Writer.WriteLine(Encrypt("!!~!!00023%^%" + SaveProject_DataSet.Tables[0].Rows[AppRec_Counter]["Pitch_or_TPI"].ToString().Trim() + "//~//", "PolarissimaDwarf", "WizERPEmance", "SHA1", 10));
                        Project_Writer.WriteLine(Encrypt("!!~!!00024%^%" + SaveProject_DataSet.Tables[0].Rows[AppRec_Counter]["Pitch_TPI_Value"].ToString().Trim() + "//~//", "PolarissimaDwarf", "WizERPEmance", "SHA1", 10));
                        Project_Writer.WriteLine(Encrypt("!!~!!00025%^%" + SaveProject_DataSet.Tables[0].Rows[AppRec_Counter]["Number_of_Bolts"].ToString().Trim() + "//~//", "PolarissimaDwarf", "WizERPEmance", "SHA1", 10));
                        Project_Writer.WriteLine(Encrypt("!!~!!00026%^%" + SaveProject_DataSet.Tables[0].Rows[AppRec_Counter]["LTF"].ToString().Trim() + "//~//", "PolarissimaDwarf", "WizERPEmance", "SHA1", 10));
                        Project_Writer.WriteLine(Encrypt("!!~!!00027%^%" + SaveProject_DataSet.Tables[0].Rows[AppRec_Counter]["Bolt_to_ToolRatioId"].ToString().Trim() + "//~//", "PolarissimaDwarf", "WizERPEmance", "SHA1", 10));
                        Project_Writer.WriteLine(Encrypt("!!~!!00028%^%" + SaveProject_DataSet.Tables[0].Rows[AppRec_Counter]["Bolt_to_ToolRatio"].ToString().Trim() + "//~//", "PolarissimaDwarf", "WizERPEmance", "SHA1", 10));
                        Project_Writer.WriteLine(Encrypt("!!~!!00029%^%" + SaveProject_DataSet.Tables[0].Rows[AppRec_Counter]["ToolId"].ToString().Trim() + "//~//", "PolarissimaDwarf", "WizERPEmance", "SHA1", 10));
                        Project_Writer.WriteLine(Encrypt("!!~!!00030%^%" + SaveProject_DataSet.Tables[0].Rows[AppRec_Counter]["ModelNumber"].ToString().Trim() + "//~//", "PolarissimaDwarf", "WizERPEmance", "SHA1", 10));
                        Project_Writer.WriteLine(Encrypt("!!~!!00031%^%" + SaveProject_DataSet.Tables[0].Rows[AppRec_Counter]["ToolPressureArea"].ToString().Trim() + "//~//", "PolarissimaDwarf", "WizERPEmance", "SHA1", 10));
                        Project_Writer.WriteLine(Encrypt("!!~!!00032%^%" + SaveProject_DataSet.Tables[0].Rows[AppRec_Counter]["MaterialId"].ToString().Trim() + "//~//", "PolarissimaDwarf", "WizERPEmance", "SHA1", 10));
                        Project_Writer.WriteLine(Encrypt("!!~!!00033%^%" + SaveProject_DataSet.Tables[0].Rows[AppRec_Counter]["Material"].ToString().Trim() + "//~//", "PolarissimaDwarf", "WizERPEmance", "SHA1", 10));
                        Project_Writer.WriteLine(Encrypt("!!~!!00034%^%" + SaveProject_DataSet.Tables[0].Rows[AppRec_Counter]["BoltYield"].ToString().Trim() + "//~//", "PolarissimaDwarf", "WizERPEmance", "SHA1", 10));
                        Project_Writer.WriteLine(Encrypt("!!~!!00035%^%" + SaveProject_DataSet.Tables[0].Rows[AppRec_Counter]["BoltTensileStressArea"].ToString().Trim() + "//~//", "PolarissimaDwarf", "WizERPEmance", "SHA1", 10));
                        Project_Writer.WriteLine(Encrypt("!!~!!00036%^%" + SaveProject_DataSet.Tables[0].Rows[AppRec_Counter]["BoltLength"].ToString().Trim() + "//~//", "PolarissimaDwarf", "WizERPEmance", "SHA1", 10));
                        Project_Writer.WriteLine(Encrypt("!!~!!00037%^%" + SaveProject_DataSet.Tables[0].Rows[AppRec_Counter]["BoltStressBase"].ToString().Trim() + "//~//", "PolarissimaDwarf", "WizERPEmance", "SHA1", 10));
                        Project_Writer.WriteLine(Encrypt("!!~!!00038%^%" + SaveProject_DataSet.Tables[0].Rows[AppRec_Counter]["BoltMinorDiameterArea"].ToString().Trim() + "//~//", "PolarissimaDwarf", "WizERPEmance", "SHA1", 10));
                        Project_Writer.WriteLine(Encrypt("!!~!!00039%^%" + SaveProject_DataSet.Tables[0].Rows[AppRec_Counter]["T1_APressureBoltStress"].ToString().Trim() + "//~//", "PolarissimaDwarf", "WizERPEmance", "SHA1", 10));
                        Project_Writer.WriteLine(Encrypt("!!~!!00040%^%" + SaveProject_DataSet.Tables[0].Rows[AppRec_Counter]["T1_APressureBoltLoad"].ToString().Trim() + "//~//", "PolarissimaDwarf", "WizERPEmance", "SHA1", 10));
                        Project_Writer.WriteLine(Encrypt("!!~!!00041%^%" + SaveProject_DataSet.Tables[0].Rows[AppRec_Counter]["T1_APressureBoltYieldPC"].ToString().Trim() + "//~//", "PolarissimaDwarf", "WizERPEmance", "SHA1", 10));
                        Project_Writer.WriteLine(Encrypt("!!~!!00042%^%" + SaveProject_DataSet.Tables[0].Rows[AppRec_Counter]["T1_BPressureBoltStress"].ToString().Trim() + "//~//", "PolarissimaDwarf", "WizERPEmance", "SHA1", 10));
                        Project_Writer.WriteLine(Encrypt("!!~!!00043%^%" + SaveProject_DataSet.Tables[0].Rows[AppRec_Counter]["T1_BPressureBoltLoad"].ToString().Trim() + "//~//", "PolarissimaDwarf", "WizERPEmance", "SHA1", 10));
                        Project_Writer.WriteLine(Encrypt("!!~!!00044%^%" + SaveProject_DataSet.Tables[0].Rows[AppRec_Counter]["T1_BPressureBoltYieldPC"].ToString().Trim() + "//~//", "PolarissimaDwarf", "WizERPEmance", "SHA1", 10));
                        Project_Writer.WriteLine(Encrypt("!!~!!00045%^%" + SaveProject_DataSet.Tables[0].Rows[AppRec_Counter]["T2_ResidualBoltStress"].ToString().Trim() + "//~//", "PolarissimaDwarf", "WizERPEmance", "SHA1", 10));
                        Project_Writer.WriteLine(Encrypt("!!~!!00046%^%" + SaveProject_DataSet.Tables[0].Rows[AppRec_Counter]["T2_ResidualBoltLoad"].ToString().Trim() + "//~//", "PolarissimaDwarf", "WizERPEmance", "SHA1", 10));
                        Project_Writer.WriteLine(Encrypt("!!~!!00047%^%" + SaveProject_DataSet.Tables[0].Rows[AppRec_Counter]["T2_ResidualBoltYieldPC"].ToString().Trim() + "//~//", "PolarissimaDwarf", "WizERPEmance", "SHA1", 10));
                        Project_Writer.WriteLine(Encrypt("!!~!!00048%^%" + SaveProject_DataSet.Tables[0].Rows[AppRec_Counter]["DetensioningBoltStress"].ToString().Trim() + "//~//", "PolarissimaDwarf", "WizERPEmance", "SHA1", 10));
                        Project_Writer.WriteLine(Encrypt("!!~!!00049%^%" + SaveProject_DataSet.Tables[0].Rows[AppRec_Counter]["DetensioningBoltLoad"].ToString().Trim() + "//~//", "PolarissimaDwarf", "WizERPEmance", "SHA1", 10));
                        Project_Writer.WriteLine(Encrypt("!!~!!00050%^%" + SaveProject_DataSet.Tables[0].Rows[AppRec_Counter]["DetensioningBoltYieldPC"].ToString().Trim() + "//~//", "PolarissimaDwarf", "WizERPEmance", "SHA1", 10));
                        Project_Writer.WriteLine(Encrypt("!!~!!00051%^%" + SaveProject_DataSet.Tables[0].Rows[AppRec_Counter]["TensionPressure_FirstPass"].ToString().Trim() + "//~//", "PolarissimaDwarf", "WizERPEmance", "SHA1", 10));
                        Project_Writer.WriteLine(Encrypt("!!~!!00052%^%" + SaveProject_DataSet.Tables[0].Rows[AppRec_Counter]["TensionPressure_SecondPass"].ToString().Trim() + "//~//", "PolarissimaDwarf", "WizERPEmance", "SHA1", 10));
                        Project_Writer.WriteLine(Encrypt("!!~!!00053%^%" + SaveProject_DataSet.Tables[0].Rows[AppRec_Counter]["TensionPressure_ThirdPass"].ToString().Trim() + "//~//", "PolarissimaDwarf", "WizERPEmance", "SHA1", 10));
                        Project_Writer.WriteLine(Encrypt("!!~!!00054%^%" + SaveProject_DataSet.Tables[0].Rows[AppRec_Counter]["CheckingPass"].ToString().Trim() + "//~//", "PolarissimaDwarf", "WizERPEmance", "SHA1", 10));
                        Project_Writer.WriteLine(Encrypt("!!~!!00055%^%" + SaveProject_DataSet.Tables[0].Rows[AppRec_Counter]["Torque"].ToString().Trim() + "//~//", "PolarissimaDwarf", "WizERPEmance", "SHA1", 10));
                        Project_Writer.WriteLine(Encrypt("!!~!!00056%^%" + SaveProject_DataSet.Tables[0].Rows[AppRec_Counter]["Coefficient_Friction"].ToString().Trim() + "//~//", "PolarissimaDwarf", "WizERPEmance", "SHA1", 10));
                        Project_Writer.WriteLine(Encrypt("!!~!!00057%^%" + SaveProject_DataSet.Tables[0].Rows[AppRec_Counter]["Bolt01"].ToString().Trim() + "//~//", "PolarissimaDwarf", "WizERPEmance", "SHA1", 10));
                        Project_Writer.WriteLine(Encrypt("!!~!!00058%^%" + SaveProject_DataSet.Tables[0].Rows[AppRec_Counter]["Bolt02"].ToString().Trim() + "//~//", "PolarissimaDwarf", "WizERPEmance", "SHA1", 10));
                        Project_Writer.WriteLine(Encrypt("!!~!!00059%^%" + SaveProject_DataSet.Tables[0].Rows[AppRec_Counter]["Bolt03"].ToString().Trim() + "//~//", "PolarissimaDwarf", "WizERPEmance", "SHA1", 10));
                        Project_Writer.WriteLine(Encrypt("!!~!!00060%^%" + SaveProject_DataSet.Tables[0].Rows[AppRec_Counter]["Bolt04"].ToString().Trim() + "//~//", "PolarissimaDwarf", "WizERPEmance", "SHA1", 10));
                        Project_Writer.WriteLine(Encrypt("!!~!!00061%^%" + SaveProject_DataSet.Tables[0].Rows[AppRec_Counter]["FirstPass_100"].ToString().Trim() + "//~//", "PolarissimaDwarf", "WizERPEmance", "SHA1", 10));
                        Project_Writer.WriteLine(Encrypt("!!~!!00062%^%" + SaveProject_DataSet.Tables[0].Rows[AppRec_Counter]["FirstPass_50"].ToString().Trim() + "//~//", "PolarissimaDwarf", "WizERPEmance", "SHA1", 10));
                        Project_Writer.WriteLine(Encrypt("!!~!!00063%^%" + SaveProject_DataSet.Tables[0].Rows[AppRec_Counter]["SecondPass_50"].ToString().Trim() + "//~//", "PolarissimaDwarf", "WizERPEmance", "SHA1", 10));
                        Project_Writer.WriteLine(Encrypt("!!~!!00064%^%" + SaveProject_DataSet.Tables[0].Rows[AppRec_Counter]["Max_Detensioning"].ToString().Trim() + "//~//", "PolarissimaDwarf", "WizERPEmance", "SHA1", 10));
                        Project_Writer.WriteLine(Encrypt("!!~!!00065%^%" + SaveProject_DataSet.Tables[0].Rows[AppRec_Counter]["ResidualBoltStress"].ToString().Trim() + "//~//", "PolarissimaDwarf", "WizERPEmance", "SHA1", 10));
                        Project_Writer.WriteLine(Encrypt("!!~!!00066%^%" + SaveProject_DataSet.Tables[0].Rows[AppRec_Counter]["Comments"].ToString().Trim() + "//~//", "PolarissimaDwarf", "WizERPEmance", "SHA1", 10));
                        Project_Writer.WriteLine(Encrypt("!!~!!00067%^%" + SaveProject_DataSet.Tables[0].Rows[AppRec_Counter]["IsDetension"].ToString().Trim() + "//~//", "PolarissimaDwarf", "WizERPEmance", "SHA1", 10));
                        Project_Writer.WriteLine(Encrypt("!!~!!00068%^%" + SaveProject_DataSet.Tables[0].Rows[AppRec_Counter]["CrossLoadingPC"].ToString().Trim() + "//~//", "PolarissimaDwarf", "WizERPEmance", "SHA1", 10));
                        Project_Writer.WriteLine(Encrypt("!!~!!00069%^%" + SaveProject_DataSet.Tables[0].Rows[AppRec_Counter]["DetensioningPC"].ToString().Trim() + "//~//", "PolarissimaDwarf", "WizERPEmance", "SHA1", 10));
                        Project_Writer.WriteLine(Encrypt("!!~!!00070%^%" + SaveProject_DataSet.Tables[0].Rows[AppRec_Counter]["T3_ResidualBoltStress"].ToString().Trim() + "//~//", "PolarissimaDwarf", "WizERPEmance", "SHA1", 10));
                        Project_Writer.WriteLine(Encrypt("!!~!!00071%^%" + SaveProject_DataSet.Tables[0].Rows[AppRec_Counter]["T3_ResidualBoltLoad"].ToString().Trim() + "//~//", "PolarissimaDwarf", "WizERPEmance", "SHA1", 10));
                        Project_Writer.WriteLine(Encrypt("!!~!!00072%^%" + SaveProject_DataSet.Tables[0].Rows[AppRec_Counter]["T3_ResidualBoltYieldPC"].ToString().Trim() + "//~//", "PolarissimaDwarf", "WizERPEmance", "SHA1", 10));
                        Project_Writer.WriteLine(Encrypt("!!~!!AppEnd " + (AppRec_Counter + 1).ToString() + "//~//", "PolarissimaDwarf", "WizERPEmance", "SHA1", 10));
                    }
                    Project_Writer.WriteLine(Encrypt("!!~!!Application_End//~//", "PolarissimaDwarf", "WizERPEmance", "SHA1", 10));
                    Project_Writer.WriteLine(Encrypt("!!~!!Copyright: Wizodyssey Software 2016; All rights reserved.//~//", "PolarissimaDwarf", "WizERPEmance", "SHA1", 10));
                    Project_Writer.WriteLine(Encrypt("!!~!!Conceived, Designed and Developed by Wizodyssey Software; Address: Unit 628 & 629, Master Mid IV, Royal Palms, Aarey Milk Colony, Goregaon (E), MUMBAI - 400065; Email: sales@wizodyssey.com; Mobile Phone: 91-9594298558; Contact: Nitin Rajurkar //~//", "PolarissimaDwarf", "WizERPEmance", "SHA1", 10));
                }
            }
            catch (Exception AppWrite_Exceprion)
            {
                MessageBox.Show(AppWrite_Exceprion.Message + " Error occurred while writing applications to file.", "Error!");
            }
            finally
            {
                Project_Writer.Close();
                Project_Writer.Dispose();
            }
        }
        public string OpenProject()
        {
            string ProjectResult = string.Empty;
            if (Convert.ToInt32(ProjectId_Label.Text.Trim()) > 0)
            {
                DialogResult IsProjectPersistence = MessageBox.Show("Do you want to save current project?", "User Confirmation!", MessageBoxButtons.YesNoCancel);
                switch (IsProjectPersistence)
                {
                    case DialogResult.Yes:
                        SaveClose_Project();
                        ProjectResult = ProjectManifest();
                        break;
                    case DialogResult.No:
                        CloseProject();
                        ProjectResult = ProjectManifest();
                        break;
                    case DialogResult.Cancel:
                        ProjectResult = "Cancel";
                        break;
                    default:
                        ProjectResult = "Cancel";
                        break;
                }
            }
            else
            {
                ProjectResult = ProjectManifest();
            }
            
            return ProjectResult;
        }

        protected string ProjectManifest()
        {
            string ManifestResult = string.Empty;
            OpenFileDialog OpenProject = new OpenFileDialog();
            OpenProject.Filter = "Project Files (*.wizbolt)|*.wizbolt";
            OpenProject.DefaultExt = "wizbolt";
            OpenProject.RestoreDirectory = true;
            OpenProject.Multiselect = false;
            DialogResult OpenResult = OpenProject.ShowDialog();

            if (OpenResult == DialogResult.OK)
            {
                ProjectOpenedFile = OpenProject.FileName;
               // DefaultStress_Button.Visible = true;
                ManifestResult = OpenProjectFile(ProjectOpenedFile);
            }
            IsLoading = false;
            return ManifestResult;
        }

        protected string OpenProjectFile(string ProjectFileWithPath)
        {
            string ManifestResult = string.Empty;             // Returns the success or failure
            string ReadStream = string.Empty;                // Reads project data after commencement of project data tag is encountered
            string ProjectValues = string.Empty;            // Project attribute values read from file one by one and assigned to ValueList
            string ValueList = string.Empty;                // Concatenated string of project attribute values separated by comma and numerals not included in single quotes as required by INSERT SQL statement.
            decimal NumericValue = 0M;                      // To separate numerals from strings by parsing project attribute values so as to include in single quote or not. Numbers are not included in single quotes.
            int Number_of_Apps = 0;                         // How many applications are there in this project. File has this value and it is read.
            string AppValues = string.Empty;                // Application values read from the file
            decimal AppNumericValues = 0M;
            string AppValueList = string.Empty;             // 

            if (File.Exists(ProjectFileWithPath))
            {
                FileWithPath_Label.Text = ProjectFileWithPath;
                SQLiteConnection ConnectDB = new SQLiteConnection("Data Source=" + BasePath + "\\WizBolt.db;New=False;");
                SQLiteCommand OpenProject_SQLiteCommand = new SQLiteCommand(ConnectDB);

                FileStream ReadProject_FileStream = new FileStream(ProjectFileWithPath, FileMode.Open);
                StreamReader Project_Reader = new StreamReader(ReadProject_FileStream);
                try
                {
                    string Raw_Header = Project_Reader.ReadLine();              // Reads header which is Wizodyssey that recognize that the file is valid format of WizBolt Utility.
                    string Header = Decrypt(Raw_Header,  "PolarissimaDwarf", "WizERPEmance", "SHA1", 10);       // Decrypts the da
                    Header = Header.Substring(5, 10);
                    if (Header == "Wizodyssey")
                    {
                        string Raw_CommenceTag = Project_Reader.ReadLine();     // Indicates commencement of project data
                        string CommenceTag = Decrypt(Raw_CommenceTag, "PolarissimaDwarf", "WizERPEmance", "SHA1", 10);
                        CommenceTag = CommenceTag.Substring(5, 13);
                        if (CommenceTag == "Project_Start")
                        {
                            while (ReadStream != "Project_End")             // Parse project data
                            {
                                ReadStream = Decrypt(Project_Reader.ReadLine(), "PolarissimaDwarf", "WizERPEmance", "SHA1", 10).Substring(5);
                                if (ReadStream.Contains("%^%"))         // Value of the project attribute
                                {
                                    ReadStream = ReadStream.Substring(7);
                                    ReadStream = ReadStream.Substring(0, (ReadStream.Length - 5));
                                }
                                else                   // Tag to indicate operation
                                {
                                    ReadStream = ReadStream.Substring(0, (ReadStream.Length - 5));
                                }
                                ProjectValues = ReadStream;
                                if (ValueList == string.Empty)
                                {
                                    ValueList = "'" + ProjectValues + "'";
                                }
                                else
                                {
                                    if (decimal.TryParse(ProjectValues, out NumericValue))
                                    {
                                        ValueList = ValueList + ", " + NumericValue.ToString();
                                    }
                                    else
                                    {
                                        ValueList = ValueList + ", '" + ProjectValues + "'";
                                    }
                                }
                            }
                            ValueList = ValueList.Substring(0, (ValueList.Length - 15));
                            // Insert project in ProjectReport table
                            ConnectDB.Open();

                            OpenProject_SQLiteCommand.CommandText = "INSERT INTO ProjectReport (ProjectReportId, CustomerId, CustomerName, CustomerLocation, ProjectName, "
                                                                       + " ProjectReference, StartDateDay, StartDateMonth, StartDateYear, "
                                                                       + " EndDateDay, EndDateMonth, EndDateYear, "
                                                                       + " EngineerName,  Notes, SummaryNotes, "
                                                                       + " TensionerToolSeriesId, Series, CrossLoading_PC, Detensioning_PC, Coefficient_of_Friction, StressValue_Base"
                                                                       + ")"
                                                                       + " VALUES (1, 1, " + ValueList + ");";

                            int ProjectResult = OpenProject_SQLiteCommand.ExecuteNonQuery();
                            if (ProjectResult > 0)
                            {
                                ReadStream = Decrypt(Project_Reader.ReadLine(), "PolarissimaDwarf", "WizERPEmance", "SHA1", 10).Substring(5, 17);
                                if (ReadStream == "Application_Start")          // Parse application data
                                {
                                    string Number_of_Applications = Decrypt(Project_Reader.ReadLine(), "PolarissimaDwarf", "WizERPEmance", "SHA1", 10).Substring(5);
                                    if (Number_of_Applications.Length > 5)
                                    {
                                        Number_of_Applications = Number_of_Applications.Substring(0, (Number_of_Applications.Length - 5));
                                        if (Int32.TryParse(Number_of_Applications, out Number_of_Apps))
                                        {
                                            for (int AppCounter = 0; AppCounter < Number_of_Apps; AppCounter++)
                                            {
                                                ReadStream = Decrypt(Project_Reader.ReadLine(), "PolarissimaDwarf", "WizERPEmance", "SHA1", 10).Substring(5);
                                                while (AppValues != "AppEnd")            // Get each application separately
                                                {
                                                    AppValues = Decrypt(Project_Reader.ReadLine(), "PolarissimaDwarf", "WizERPEmance", "SHA1", 10).Substring(5);
                                                    if (AppValues.Contains("%^%"))      // Value of application attributes
                                                    {
                                                        AppValues = AppValues.Substring(8);
                                                        AppValues = AppValues.Substring(0, (AppValues.Length - 5));
                                                    }
                                                    else                              // Tag to indicate operation
                                                    {
                                                        AppValues = AppValues.Substring(0, 6);
                                                    }
                                                    // Concate values required for INSERT statement of SQL
                                                    if (AppValueList == string.Empty)
                                                    {
                                                        AppValueList = AppValues;
                                                    }
                                                    else
                                                    {
                                                        if (decimal.TryParse(AppValues, out AppNumericValues))
                                                        {
                                                            AppValueList = AppValueList + ", " + AppNumericValues.ToString();
                                                        }
                                                        else
                                                        {
                                                            AppValueList = AppValueList + ", '" + AppValues + "'";
                                                        }
                                                    }
                                                }
                                                AppValueList = AppValueList.Substring(0, (AppValueList.Length - 10));    // Remove tag 'AppEnd'
                                                AppValues = string.Empty;

                                                // Insert the application in ProjectDetailedReport table
                                                OpenProject_SQLiteCommand.CommandText = "INSERT INTO ProjectDetailedReport (ProjectReportId, ApplicationId, UnitSystem, JointId, PrimaryStandardId,  "
                                                                                      + "   Specification, PrimaryFlangeRatingId, FlangeRating, Flange1_TypeId,  "
                                                                                      + "   Flange1_Abbreviation, Flange1_ClampLength, Flange2_TypeId, Flange2_Abbreviation,  "
                                                                                      + "   Flange2_ClampLength, GasketId, Gasket, GasketGap,  "
                                                                                      + "   SpacerId, Spacer, SpacerThickness, TotalClampLength, PrimaryFlangeBoltId,  "
                                                                                      + "   BoltThread, Pitch_or_TPI, Pitch_TPI_Value, Number_of_Bolts, LTF, "
                                                                                      + "   Bolt_to_ToolRatioId, Bolt_to_ToolRatio, ToolId, ModelNumber,  "
                                                                                      + "   ToolPressureArea, MaterialId, Material, BoltYield, BoltTensileStressArea, "
                                                                                      + "   BoltLength, BoltStressBase, BoltMinorDiameterArea, T1_APressureBoltStress, T1_APressureBoltLoad,  "
                                                                                      + "   T1_APressureBoltYieldPC, T1_BPressureBoltStress, T1_BPressureBoltLoad, T1_BPressureBoltYieldPC,   "
                                                                                      + "   T2_ResidualBoltStress, T2_ResidualBoltLoad, T2_ResidualBoltYieldPC, DetensioningBoltStress, DetensioningBoltLoad,	"
                                                                                      + "   DetensioningBoltYieldPC, TensionPressure_FirstPass, TensionPressure_SecondPass, TensionPressure_ThirdPass, CheckingPass,   "
                                                                                      + "   Torque, Coefficient_Friction, Bolt01, Bolt02, Bolt03, Bolt04, FirstPass_100, FirstPass_50, SecondPass_50,  "
                                                                                      + "   Max_Detensioning, ResidualBoltStress, Comments, IsDetension, CrossLoadingPC, DetensioningPC, "
                                                                                      + "   T3_ResidualBoltStress, T3_ResidualBoltLoad, T3_ResidualBoltYieldPC)  "
                                                                                      + " VALUES (1," + AppValueList + ");";
                                                // Execute to insert data in table
                                                int AppResult = OpenProject_SQLiteCommand.ExecuteNonQuery();
                                                AppValueList = string.Empty;
                                            }
                                        }
                                    }
                                }
                            }
                            ConnectDB.Close();
                        }
                    }
                }
                catch (Exception Read_Exceprion)
                {
                    MessageBox.Show(Read_Exceprion.Message + " Error occurred while reading file.", "Error!");
                    ManifestResult = "Failed";
                }
                finally
                {
                    OpenProject_SQLiteCommand.Dispose();
                    ConnectDB.Dispose();
                    ReadProject_FileStream.Close();
                    ReadProject_FileStream.Dispose();
                }
                try
                {
                    // Project along with its applications are read and restored to database. Now display it to user.
                    ProjectId_Label.Text = "1";
                    CustomerId_Label.Text = "1";
                    AppId_Label.Text = "1";

                    SQLiteConnection ConnectOpenedProject = new SQLiteConnection("Data Source=WizBolt.db;New=False;");
                        SQLiteCommand SelectOpenedProject_SQLiteCommand = new SQLiteCommand(ConnectOpenedProject);
                        ConnectOpenedProject.Open();
                        SelectOpenedProject_SQLiteCommand.CommandText = "SELECT 	P.CustomerId, P.CustomerName, P.CustomerLocation, P.ProjectReportId, P.ProjectName, P.ProjectReference, "
                                                                        + " P.StartDateDay, P.StartDateMonth, P.StartDateYear, P.EndDateDay, P.EndDateMonth, P.EndDateYear, "
                                                                        + " P.EngineerName, P.Notes, P.SummaryNotes, TensionerToolSeriesId, Series, CrossLoading_PC, Detensioning_PC, "
                                                                        + " Coefficient_of_Friction, StressValue_Base"
                                                                        + " 	FROM ProjectReport	P   "
                                                                        + "             WHERE P.ProjectReportId > 0 AND P.ProjectReportId = " + ProjectId_Label.Text + "; ";
                        DataSet OpenedProject_DataSet = new DataSet();
                        SQLiteDataAdapter OpenedProject_DataAdapter = new SQLiteDataAdapter(SelectOpenedProject_SQLiteCommand);
                        int OpenedProjectResult = OpenedProject_DataAdapter.Fill(OpenedProject_DataSet);
                        if (OpenedProjectResult > 0)
                        {
                            IsLoading = true;
                            CustomerName_Label.Text = OpenedProject_DataSet.Tables[0].Rows[0]["CustomerName"].ToString();
                            LocationName_Label.Text = OpenedProject_DataSet.Tables[0].Rows[0]["CustomerLocation"].ToString();
                            ProjectName_Label.Text = OpenedProject_DataSet.Tables[0].Rows[0]["ProjectName"].ToString();
                            Reference_Label.Text = OpenedProject_DataSet.Tables[0].Rows[0]["ProjectReference"].ToString();
                            EnteredProjectDate_label.Text = OpenedProject_DataSet.Tables[0].Rows[0]["StartDateDay"].ToString() + "-"
                                                            + OpenedProject_DataSet.Tables[0].Rows[0]["StartDateMonth"].ToString() + "-"
                                                            + OpenedProject_DataSet.Tables[0].Rows[0]["StartDateYear"].ToString();

                            if ((OpenedProject_DataSet.Tables[0].Rows[0]["EndDateDay"].ToString().Trim() == "00") || (OpenedProject_DataSet.Tables[0].Rows[0]["EndDateDay"].ToString().Trim() == "0"))
                            {
                                ProjectEndDate_Label.Text = "Not Entered";
                            }
                            else
                            {
                                ProjectEndDate_Label.Text = OpenedProject_DataSet.Tables[0].Rows[0]["EndDateDay"].ToString().Trim() + "-"
                                                            + OpenedProject_DataSet.Tables[0].Rows[0]["EndDateMonth"].ToString().Trim() + "-"
                                                            + OpenedProject_DataSet.Tables[0].Rows[0]["EndDateYear"].ToString().Trim();

                            }
                            ToolSeriesId_Label.Text = OpenedProject_DataSet.Tables[0].Rows[0]["TensionerToolSeriesId"].ToString().Trim();
                            ToolSeries_Label.Text = OpenedProject_DataSet.Tables[0].Rows[0]["Series"].ToString().Trim();
                            Engineer_Label.Text = OpenedProject_DataSet.Tables[0].Rows[0]["EngineerName"].ToString().Trim();
                            Notes_RichTextBox.Text = OpenedProject_DataSet.Tables[0].Rows[0]["Notes"].ToString().Trim();
                            Summary_Label.Text = OpenedProject_DataSet.Tables[0].Rows[0]["SummaryNotes"].ToString().Trim();
                            CrossLoading_TextBox.Text = OpenedProject_DataSet.Tables[0].Rows[0]["CrossLoading_PC"].ToString().Trim();
                            Detensioning_TextBox.Text = OpenedProject_DataSet.Tables[0].Rows[0]["Detensioning_PC"].ToString().Trim();
                            CoefficientValue_Label.Text = OpenedProject_DataSet.Tables[0].Rows[0]["Coefficient_of_Friction"].ToString().Trim();
                            Stress_Label.Text = OpenedProject_DataSet.Tables[0].Rows[0]["StressValue_Base"].ToString().Trim();

                            CustomerName_Label.Visible = true;
                            LocationName_Label.Visible = true;
                            ProjectName_Label.Visible = true;
                            Reference_Label.Visible = true;
                            ProjectDate_Label.Visible = true;
                            EnteredProjectDate_label.Visible = true;
                            ProjectEndDateTitle_Label.Visible = true;
                            ProjectEndDate_Label.Visible = true;
                            ToolSeries_Label.Visible = true;
                            Engineer_Label.Visible = true;

                            // Application List displayed in Data Grid
                            SelectOpenedProject_SQLiteCommand.CommandText = "SELECT  PR.CustomerId, PR.ProjectReportId, ApplicationId, JointId, "
                                                                    + "   FlangeRating, Flange1_Abbreviation, Flange2_Abbreviation, "
                                                                    + "   (BoltThread || ' x ' ||	Pitch_TPI_Value) AS BoltSize, Number_of_Bolts, Material "
                                                                    + "   FROM ProjectReport   PR"
                                                                    + "   	INNER JOIN ProjectDetailedReport     PD 		ON		PD.ProjectReportId = PR.ProjectReportId "
                                                                    + "   	WHERE PR.ProjectReportId > 0  AND PR.ProjectReportId = " + ProjectId_Label.Text + ";";
                            DataSet AppGrid_DataSet = new DataSet();
                            SQLiteDataAdapter AppGrid_DataAdapter = new SQLiteDataAdapter(SelectOpenedProject_SQLiteCommand);

                            // Setting DataGridView
                            int AppGridResult = AppGrid_DataAdapter.Fill(AppGrid_DataSet);
                            if (AppGridResult > 0)
                            {
                                if (Project_DataGridView.DataSource != null)
                                {
                                    Project_DataGridView.DataSource = null;
                                }
                                Project_DataGridView.AutoGenerateColumns = false;
                                Project_DataGridView.ColumnCount = 10;
                                Project_DataGridView.Columns[0].Visible = false;
                                Project_DataGridView.Columns[0].Name = "Customer_Id";
                                Project_DataGridView.Columns[0].HeaderText = "Customer Id";
                                Project_DataGridView.Columns[0].DataPropertyName = "CustomerId";
                                Project_DataGridView.Columns[1].Visible = false;
                                Project_DataGridView.Columns[1].Name = "ProjectReport_Id";
                                Project_DataGridView.Columns[1].HeaderText = "Project Report Id";
                                Project_DataGridView.Columns[1].DataPropertyName = "ProjectReportId";
                                Project_DataGridView.Columns[2].Visible = false;
                                Project_DataGridView.Columns[2].Name = "Application_Id";
                                Project_DataGridView.Columns[2].HeaderText = "Application Id";
                                Project_DataGridView.Columns[2].DataPropertyName = "ApplicationId";
                                Project_DataGridView.Columns[3].Visible = true;
                                Project_DataGridView.Columns[3].Name = "Joint_Id";
                                Project_DataGridView.Columns[3].HeaderText = "ID";
                                Project_DataGridView.Columns[3].DataPropertyName = "JointId";
                                Project_DataGridView.Columns[3].Width = 120;
                                Project_DataGridView.Columns[4].Visible = true;
                                Project_DataGridView.Columns[4].Name = "Flange_Rating";
                                Project_DataGridView.Columns[4].HeaderText = "Rating";
                                Project_DataGridView.Columns[4].DataPropertyName = "FlangeRating";
                                Project_DataGridView.Columns[4].Width = 180;
                                Project_DataGridView.Columns[5].Visible = true;
                                Project_DataGridView.Columns[5].Name = "Flange1";
                                Project_DataGridView.Columns[5].HeaderText = "Flange 1";
                                Project_DataGridView.Columns[5].DataPropertyName = "Flange1_Abbreviation";
                                Project_DataGridView.Columns[6].Visible = true;
                                Project_DataGridView.Columns[6].Name = "Flange2";
                                Project_DataGridView.Columns[6].HeaderText = "Flange 2";
                                Project_DataGridView.Columns[6].DataPropertyName = "Flange2_Abbreviation";
                                Project_DataGridView.Columns[7].Visible = true;
                                Project_DataGridView.Columns[7].Name = "Bolt_Diameter";
                                Project_DataGridView.Columns[7].HeaderText = "Bolt Size";
                                Project_DataGridView.Columns[7].DataPropertyName = "BoltSize";
                                Project_DataGridView.Columns[8].Visible = true;
                                Project_DataGridView.Columns[8].Name = "Bolts";
                                Project_DataGridView.Columns[8].HeaderText = "Bolts";
                                Project_DataGridView.Columns[8].DataPropertyName = "Number_of_Bolts";
                                Project_DataGridView.Columns[9].Visible = true;
                                Project_DataGridView.Columns[9].Name = "Bolt_Material";
                                Project_DataGridView.Columns[9].HeaderText = "Bolt Material";
                                Project_DataGridView.Columns[9].DataPropertyName = "Material";
                                Project_DataGridView.Columns[9].Width = 160;
                                Project_DataGridView.DataSource = AppGrid_DataSet.Tables[0];
                            }
                        }
                    ConnectOpenedProject.Close();
                    SelectOpenedProject_SQLiteCommand.Dispose();
                    ConnectOpenedProject.Dispose();
                    ManifestResult = "Successful";
                    IsLoading = false;
                }
                catch (Exception Display_Exceprion)
                {
                    MessageBox.Show(Display_Exceprion.Message + " Error occurred while displaying project with applications that were read from file.", "Error!");
                    ManifestResult = "Failed";
                }
                
            }
            else
            {
                MessageBox.Show("File cannot be opened. Either file does not exists or you do not have right to open it.", "Fatal Error!!!");
                ManifestResult = "Failed";
            }
            return ManifestResult;
        }

    

        private void Project_DataGridView_SelectionChanged(object sender, EventArgs e)
        {
           // DataGridView dgv = sender as DataGridView;
            // Alternate way
            DataGridView  SelectedApp_DataGridView = (DataGridView)sender;
            if (SelectedApp_DataGridView != null && SelectedApp_DataGridView.SelectedRows.Count > 0)
            {
                DataGridViewRow SelectedRow = SelectedApp_DataGridView.SelectedRows[0];
                if (SelectedRow != null)
                {
                   AppId_Label.Text  = SelectedRow.Cells[2].Value.ToString().Trim();
                   SetApplication();
                }
            }
            
        }

        public void UpdateApplication()
        {
            // Local variables
            string App_Id = AppId_Label.Text;
            if ((ProjectId_Label.Text.Trim() != "") && (ProjectId_Label.Text.Trim().Length > 0) && (App_Id.Trim() != "") && (App_Id.Trim().Length > 0))
            {
                if (Convert.ToInt32(ProjectId_Label.Text.Trim()) > 0)
                {
                    // Data values from drop down list
                    string SpecificationValue = string.Empty;
                    string Flange_Rating = string.Empty;
                    string Flange1_AbbreviationValue = string.Empty;
                    string Flange2_AbbreviationValue = string.Empty;
                    string Bolt_Thread = string.Empty;
                    string Model_Number = string.Empty;
                    string Material_Value = string.Empty;
                    string Pitch_TPI_Title = string.Empty;
                    string BoltStressBase = string.Empty;
                    string Gasket_Id = string.Empty;
                    string Gasket_Name = string.Empty;
                    string Flange2_Id = string.Empty;
                    string Spacer_Id = string.Empty;
                    string Spacer_Name = string.Empty;
                    string BoltTool_RatioId = string.Empty;
                    string BoltTool_Ratio = string.Empty;
                    string Tool_Id = string.Empty;
                    string BoltDiameter_Id = string.Empty;
                    string Material_Id = string.Empty;
                    string FlangeTop_Thickness = string.Empty;
                    string FlangeBottom_Thickness = string.Empty;
                    string Gasket_Gap = string.Empty;
                    string Spacer_Thickness = string.Empty;
                    string FirstPass_100 = string.Empty;
                    string FirstPass_50 = string.Empty;
                    string SecondPass_50 = string.Empty;
                    string Max_Detensioning = string.Empty;
                    string IsDetensioned = string.Empty;
                    string Commented = string.Empty;
                    string UnitSystem = string.Empty;
                    string TorqueValue = string.Empty;

                    // Assign values of drop down list as this create problem
                    //((DataRowView)Specification_ComboBox.SelectedItem)["Standard"].ToString();
                    SpecificationValue = ((DataRowView)Specification_ComboBox.SelectedItem)["Standard"].ToString();
                    Flange_Rating = ((DataRowView)Rating_ComboBox.SelectedItem)["AbbreviatedRating"].ToString();
                    Flange1_AbbreviationValue = ((DataRowView)Flange1Config_ComboBox.SelectedItem)["FlangeAbbreviation"].ToString();

                    if ((Comment_TextBox.Text.Length == 0) && (Comment_TextBox.Text == ""))
                    {
                        Commented = "NULL";
                    }
                    else
                    {
                        Commented = Comment_TextBox.Text;
                    }
                    if (DeTension_CheckBox.Checked)
                    {
                        IsDetensioned = "Yes";
                    }
                    else
                    {
                        IsDetensioned = "No";
                    }

                    if (UnitSystem_Button.Text == "SAE")
                    {
                        UnitSystem = "SAE";

                        if (Clamp1_TextBox.Text.Length > 0)
                        {
                            FlangeTop_Thickness = (Convert.ToDecimal(Clamp1_TextBox.Text) * 25.4M).ToString();
                        }
                        else
                        {
                            FlangeTop_Thickness = "0";
                        }
                        if (GasketGap_TextBox.Text.Length > 0)
                        {
                            Gasket_Gap = (Convert.ToDecimal(GasketGap_TextBox.Text) * 25.4M).ToString();
                        }
                        else
                        {
                            Gasket_Gap = "0";
                        }
                        if (CustomSpacer_TextBox.Text.Length > 0)
                        {
                            Spacer_Thickness = (Convert.ToDecimal(CustomSpacer_TextBox.Text) * 25.4M).ToString();
                        }
                        else
                        {
                            Spacer_Thickness = "0";
                        }
                    }
                    else
                    {
                        UnitSystem = "S. I. Unit";
                        if (Clamp1_TextBox.Text.Length > 0)
                        {
                            FlangeTop_Thickness = Clamp1_TextBox.Text;
                        }
                        else
                        {
                            FlangeTop_Thickness = "0";
                        }
                        if (GasketGap_TextBox.Text.Length > 0)
                        {
                            Gasket_Gap = GasketGap_TextBox.Text;
                        }
                        else
                        {
                            Gasket_Gap = "0";
                        }
                        if (CustomSpacer_TextBox.Text.Length > 0)
                        {
                            Spacer_Thickness = CustomSpacer_TextBox.Text;
                        }
                        else
                        {
                            Spacer_Thickness = "0";
                        }
                    }

                    if (Flange2Config_ComboBox.SelectedItem != null)
                    {
                        Flange2_Id = Flange2Config_ComboBox.SelectedValue.ToString();
                        Flange2_AbbreviationValue = ((DataRowView)Flange2Config_ComboBox.SelectedItem)["FlangeAbbreviation"].ToString();
                        if (UnitSystem_Button.Text == "SAE")
                        {
                            FlangeBottom_Thickness = (Convert.ToDecimal(Clamp2_TextBox.Text) * 25.4M).ToString();
                        }
                        else
                        {
                            FlangeBottom_Thickness = Clamp2_TextBox.Text;
                        }
                    }
                    else
                    {
                        Flange2Config_ComboBox.SelectedValue = "2";
                        Flange2Config_ComboBox.SelectedItem = "WN-RF";
                        Flange2Config_ComboBox.SelectedIndex = 0;
                        Flange2_Id = "2";
                        Flange2_AbbreviationValue = "WN-RF";
                        FlangeBottom_Thickness = "25.4";
                    }
                    if (BoltThread_ComboBox.SelectedItem != null)
                    {
                        BoltDiameter_Id = BoltThread_ComboBox.SelectedValue.ToString();
                        Bolt_Thread = ((DataRowView)BoltThread_ComboBox.SelectedItem)["BoltThread"].ToString();
                    }
                    else
                    {
                        Bolt_Thread = "3/4\"";
                        BoltDiameter_Id = "1";
                    }
                    if (TensionerTool_ComboBox.SelectedItem != null)
                    {
                        Model_Number = ((DataRowView)TensionerTool_ComboBox.SelectedItem)["ModelNumber"].ToString();
                        Tool_Id = TensionerTool_ComboBox.SelectedValue.ToString();
                    }
                    else
                    {
                        Tool_Id = "";
                        Model_Number = "";
                    }
                    if (BoltMaterial_ComboBox.SelectedItem != null)
                    {
                        Material_Id = BoltMaterial_ComboBox.SelectedValue.ToString();
                        Material_Value = ((DataRowView)BoltMaterial_ComboBox.SelectedItem)["Material"].ToString();
                    }
                    else
                    {
                        Material_Value = "ASTM A193 - B7";
                        Material_Id = "3";
                    }
                    if (Gasket_ComboBox.SelectedItem.ToString() != null)
                    {
                        Gasket_Id = (Gasket_ComboBox.SelectedIndex + 1).ToString();
                        Gasket_Name = Gasket_ComboBox.SelectedItem.ToString();
                    }
                    else
                    {
                        Gasket_ComboBox.SelectedItem = "Seal Ring";
                        Gasket_ComboBox.SelectedValue = "1";
                        Gasket_Name = "Seal Ring";
                        Gasket_Id = "1";
                    }

                    if (BoltTool_ComboBox.SelectedItem != null)
                    {
                        BoltTool_RatioId = (BoltTool_ComboBox.SelectedIndex + 1).ToString();
                        BoltTool_Ratio = BoltTool_ComboBox.SelectedItem.ToString();
                    }
                    else
                    {
                        BoltTool_Ratio = "100%";
                        BoltTool_ComboBox.SelectedItem = "100%";
                        BoltTool_ComboBox.SelectedValue = "1";
                        BoltTool_RatioId = "1";
                    }
                    if (Bolt_Thread.Substring(0, 1) == "M")
                    {
                        Pitch_TPI_Title = "Pitch";
                    }
                    else
                    {
                        Pitch_TPI_Title = "TPI";
                    }
                    if (TensileStressArea_RadioButton.Checked)
                    {
                        BoltStressBase = "Tensile Stress Area";
                    }
                    else
                    {
                        BoltStressBase = "Minor Diameter Area";
                    }

                    // Treating null values. Only numberic nulls need treatment as string values has single quotes i.e. ''
                    if (Spacer_ComboBox.SelectedItem != null)
                    {
                        Spacer_Id = (Spacer_ComboBox.SelectedIndex + 1).ToString();
                        Spacer_Name = Spacer_ComboBox.SelectedItem.ToString();
                    }
                    else
                    {
                        Spacer_Id = "NULL";
                        Spacer_Name = "";
                    }
                    if ((CustomSpacer_TextBox.Text.Length == 0) && (CustomSpacer_TextBox.Text == ""))
                    {
                        CustomSpacer_TextBox.Text = "NULL";
                    }
                    if ((T1ABoltStress_SI_Label.Text.Length == 0) && (T1ABoltStress_SI_Label.Text == ""))
                    {
                        T1ABoltStress_SI_Label.Text = "NULL";
                    }
                    if ((T1ABoltLoad_SI_Label.Text.Length == 0) && (T1ABoltLoad_SI_Label.Text == ""))
                    {
                        T1ABoltLoad_SI_Label.Text = "NULL";
                    }

                    if ((T1ABoltYield_SI_Label.Text.Length == 0) && (T1ABoltYield_SI_Label.Text == ""))
                    {
                        T1ABoltYield_SI_Label.Text = "NULL";
                    }
                    if ((T1BBoltStress_SI_Label.Text.Length == 0) && (T1BBoltStress_SI_Label.Text == ""))
                    {
                        T1BBoltStress_SI_Label.Text = "NULL";
                    }
                    if ((T1BBoltLoad_SI_Label.Text.Length == 0) && (T1BBoltLoad_SI_Label.Text == ""))
                    {
                        T1BBoltLoad_SI_Label.Text = "NULL";
                    }
                    if ((T1BBoltYield_Per_Label.Text.Length == 0) && (T1BBoltYield_Per_Label.Text == ""))
                    {
                        T1BBoltYield_Per_Label.Text = "NULL";
                    }

                    if ((T2RBoltStress_SI_Label.Text.Length == 0) && (T2RBoltStress_SI_Label.Text == ""))
                    {
                        T2RBoltStress_SI_Label.Text = "NULL";
                    }
                    if ((T2RBoltLoad_SI_Label.Text.Length == 0) && (T2RBoltLoad_SI_Label.Text == ""))
                    {
                        T2RBoltLoad_SI_Label.Text = "NULL";
                    }
                    if ((T2RBoltYield_Per_Label.Text.Length == 0) && (T2RBoltYield_Per_Label.Text == ""))
                    {
                        T2RBoltYield_Per_Label.Text = "NULL";
                    }
                    if ((T3RBoltStress_SI_Label.Text.Length == 0) && (T3RBoltStress_SI_Label.Text == ""))
                    {
                        T3RBoltStress_SI_Label.Text = "NULL";
                    }
                    if ((T3RBoltLoad_SI_Label.Text.Length == 0) && (T3RBoltLoad_SI_Label.Text == ""))
                    {
                        T3RBoltLoad_SI_Label.Text = "NULL";
                    }
                    if ((T3RBoltYield_Per_Label.Text.Length == 0) && (T3RBoltYield_Per_Label.Text == ""))
                    {
                        T3RBoltYield_Per_Label.Text = "NULL";
                    }
                    if ((DetenBoltStress_SI_Label.Text.Length == 0) && (DetenBoltStress_SI_Label.Text == ""))
                    {
                        DetenBoltStress_SI_Label.Text = "NULL";
                    }
                    if ((DetenBoltLoad_SI_Label.Text.Length == 0) && (DetenBoltLoad_SI_Label.Text == ""))
                    {
                        DetenBoltLoad_SI_Label.Text = "NULL";
                    }

                    if ((DetenBoltYield_Per_Label.Text.Length == 0) && (DetenBoltYield_Per_Label.Text == ""))
                    {
                        DetenBoltYield_Per_Label.Text = "NULL";
                    }
                    if ((Pass1AppliedPressure_SequenceTabValue_SI_Label.Text.Length == 0) && (Pass1AppliedPressure_SequenceTabValue_SI_Label.Text == ""))
                    {
                        Pass1AppliedPressure_SequenceTabValue_SI_Label.Text = "NULL";
                    }
                    if ((Pass2AppliedPressure_SequenceTabValue_SI_Label.Text.Length == 0) && (Pass2AppliedPressure_SequenceTabValue_SI_Label.Text == ""))
                    {
                        Pass2AppliedPressure_SequenceTabValue_SI_Label.Text = "NULL";
                    }
                    if ((Pass3AppliedPressure_SequenceTabValue_SI_Label.Text.Length == 0) && (Pass3AppliedPressure_SequenceTabValue_SI_Label.Text == ""))
                    {
                        Pass3AppliedPressure_SequenceTabValue_SI_Label.Text = "NULL";
                    }
                    if ((CheckPass1AppliedPressure_SequenceTabValue_SI_Label.Text.Length == 0) && (CheckPass1AppliedPressure_SequenceTabValue_SI_Label.Text == ""))
                    {
                        CheckPass1AppliedPressure_SequenceTabValue_SI_Label.Text = "NULL";
                    }
                    //if (BoltTool_Ratio == "100%")
                    //{
                    //    FirstPass_100 = PressureB_SI_Label.Text.Substring(0, PressureB_SI_Label.Text.Length - 4);
                    //    Max_Detensioning = PressureB_SI_Label.Text.Substring(0, PressureB_SI_Label.Text.Length - 4);
                    //}
                    //else
                    //{
                    //    FirstPass_50 = PressureA_SI_Label.Text.Substring(0, PressureA_SI_Label.Text.Length - 4);
                    //    SecondPass_50 = PressureB_SI_Label.Text.Substring(0, PressureB_SI_Label.Text.Length - 4);
                    //    Max_Detensioning = PressureB_SI_Label.Text.Substring(0, PressureB_SI_Label.Text.Length - 4);
                    //}
                    //if ((FirstPass_100.Length == 0) && (FirstPass_100 == string.Empty))
                    //{
                    //    FirstPass_100 = "NULL";
                    //}
                    //if ((FirstPass_50.Length == 0) && (FirstPass_50 == string.Empty))
                    //{
                    //    FirstPass_50 = "NULL";
                    //}
                    //if ((SecondPass_50.Length == 0) && (SecondPass_50 == string.Empty))
                    //{
                    //    SecondPass_50 = "NULL";
                    //}
                    //if ((Max_Detensioning.Length == 0) && (Max_Detensioning == string.Empty))
                    //{
                    //    Max_Detensioning = "NULL";
                    //}

                    int Bolt_ToolId = BoltTool_ComboBox.SelectedIndex;
                    switch (Bolt_ToolId)
                    {
                        case 0:
                            if ((PressureB_SI_Label.Text.Trim() != "") && (PressureB_SI_Label.Text.Trim().Length > 4))
                            {
                                FirstPass_100 = PressureB_SI_Label.Text.Substring(0, PressureB_SI_Label.Text.Length - 4);
                                Max_Detensioning = FirstPass_100;
                            }
                            else
                            {
                                FirstPass_100 = "NULL";
                                Max_Detensioning = "NULL";
                            }

                            FirstPass_50 = "NULL";
                            SecondPass_50 = "NULL";

                            break;
                        case 1:
                            FirstPass_100 = "NULL";
                            if ((PressureA_SI_Label.Text.Trim() != "") && (PressureA_SI_Label.Text.Trim().Length > 4))
                            {
                                FirstPass_50 = PressureA_SI_Label.Text.Substring(0, PressureA_SI_Label.Text.Length - 4);
                            }
                            else
                            {
                                FirstPass_50 = "NULL";
                            }
                            if ((PressureB_SI_Label.Text.Trim() != "") && (PressureB_SI_Label.Text.Trim().Length > 4))
                            {
                                SecondPass_50 = PressureB_SI_Label.Text.Substring(0, PressureB_SI_Label.Text.Length - 4);
                                Max_Detensioning = SecondPass_50;
                            }
                            else
                            {
                                SecondPass_50 = "NULL";
                                Max_Detensioning = SecondPass_50;
                            }

                            break;
                        case 2:
                            FirstPass_100 = "NULL";
                            if ((PressureA_SI_Label.Text.Trim() != "") && (PressureA_SI_Label.Text.Trim().Length > 4))
                            {
                                FirstPass_50 = PressureA_SI_Label.Text.Substring(0, PressureA_SI_Label.Text.Length - 4);
                            }
                            else
                            {
                                FirstPass_50 = "NULL";
                            }
                            if ((PressureB_SI_Label.Text.Trim() != "") && (PressureB_SI_Label.Text.Trim().Length > 4))
                            {
                                SecondPass_50 = PressureB_SI_Label.Text.Substring(0, PressureB_SI_Label.Text.Length - 4);
                                Max_Detensioning = SecondPass_50;
                            }
                            else
                            {
                                SecondPass_50 = "NULL";
                                Max_Detensioning = SecondPass_50;
                            }

                            break;
                        case 3:
                            FirstPass_100 = "NULL";
                            FirstPass_50 = "NULL";
                            SecondPass_50 = "NULL";
                            Max_Detensioning = SecondPass_50;
                            break;
                        default:
                            FirstPass_100 = "NULL";
                            FirstPass_50 = "NULL";
                            SecondPass_50 = "NULL";
                            Max_Detensioning = SecondPass_50;
                            break;
                    }
                    TorqueValue = Torque_SI_Label.Text.Substring(0, Torque_SI_Label.Text.Length - 4);
                    if ((TorqueValue == "") && (TorqueValue == string.Empty) && (TorqueValue.Length == 0))
                    {
                        TorqueValue = "NULL";
                    }
                    else
                    {
                        try
                        {
                            decimal TorqueTest = Convert.ToDecimal(TorqueValue);
                        }
                        catch (Exception Torque_Exception)
                        {
                            MessageBox.Show(Torque_Exception.Message + " torque value is wrong. Only numerals allowed.");
                        }
                    }

                    SQLiteConnection ConnectApplication = new SQLiteConnection("Data Source=" + BasePath + "\\WizBolt.db;New=False;");
                    SQLiteCommand UpdateApplication_SQLiteCommand = new SQLiteCommand(ConnectApplication);
                    ConnectApplication.Open();
                    try
                    {
                        UpdateApplication_SQLiteCommand.CommandText = "    Update ProjectDetailedReport  SET  UnitSystem = '" + UnitSystem + "', JointId = '" + JointId_TextBox.Text + "', "
                                                   + "   PrimaryStandardId = '" + Specification_ComboBox.SelectedValue.ToString() + "', " + " Specification = '" + SpecificationValue + "', "
                                                   + "   PrimaryFlangeRatingId = '" + Rating_ComboBox.SelectedValue.ToString() + "', FlangeRating = '" + Flange_Rating + "', Flange1_TypeId = " + Flange1Config_ComboBox.SelectedValue.ToString() + ",  "
                                                   + "   Flange1_Abbreviation = '" + Flange1_AbbreviationValue + "', Flange1_ClampLength = " + FlangeTop_Thickness + ", Flange2_TypeId = " + Flange2_Id + ", Flange2_Abbreviation = '" + Flange2_AbbreviationValue + "',  "
                                                   + "   Flange2_ClampLength = " + FlangeBottom_Thickness + ", GasketId = " + Gasket_Id + ", Gasket = '" + Gasket_Name + "', GasketGap = " + Gasket_Gap + ",  "
                                                   + "   SpacerId = " + Spacer_Id + ", Spacer = '" + Spacer_Name + "', SpacerThickness = " + Spacer_Thickness + ", TotalClampLength = " + ClampLengthInfo_mm_Label.Text.Substring(0, ClampLengthInfo_mm_Label.Text.Length - 3) + ", PrimaryFlangeBoltId = '" + BoltDiameter_Id + "',  "
                                                   + "   BoltThread = '" + Bolt_Thread + "', Pitch_or_TPI = '" + Pitch_TPI_Title + "', Pitch_TPI_Value = " + Pitch_TextBox.Text + ", Number_of_Bolts = " + NumberOfBolts_TextBox.Text + ", LTF = " + LTF_Label.Text + ", "
                                                   + "   Bolt_to_ToolRatioId = " + BoltTool_RatioId + ", Bolt_to_ToolRatio = '" + BoltTool_Ratio + "', ToolId = '" + Tool_Id + "', ModelNumber = '" + Model_Number + "',  "
                                                   + "   ToolPressureArea = " + ToolPressureArea_mm_Label.Text.Substring(0, ToolPressureArea_mm_Label.Text.Length - 5) + ", MaterialId = " + Material_Id + ", Material = '" + Material_Value + "', BoltYield = " + BoltYield_SI_Label.Text.Substring(0, BoltYield_SI_Label.Text.Length - 4) + ", BoltTensileStressArea = " + TensileStressArea_SI_Label.Text.Substring(0, TensileStressArea_SI_Label.Text.Length - 5) + ", "
                                                   + "   BoltLength = " + BoltLength_SI_Label.Text.Substring(0, BoltLength_SI_Label.Text.Length - 3) + ", BoltStressBase = '" + BoltStressBase + "', BoltMinorDiameterArea = 0, T1_APressureBoltStress = " + T1ABoltStress_SI_Label.Text + ", T1_APressureBoltLoad = " + T1ABoltLoad_SI_Label.Text + ",  "
                                                   + "   T1_APressureBoltYieldPC = " + T1ABoltYield_SI_Label.Text + ", T1_BPressureBoltStress = " + T1BBoltStress_SI_Label.Text + ", T1_BPressureBoltLoad = " + T1BBoltLoad_SI_Label.Text + ", T1_BPressureBoltYieldPC = " + T1BBoltYield_Per_Label.Text + ",   "
                                                   + "   T2_ResidualBoltStress = " + T2RBoltStress_SI_Label.Text + ", T2_ResidualBoltLoad = " + T2RBoltLoad_SI_Label.Text + ", T2_ResidualBoltYieldPC = " + T2RBoltYield_Per_Label.Text + ",   T3_ResidualBoltStress = " + T3RBoltStress_SI_Label.Text + ", T3_ResidualBoltLoad = " + T3RBoltLoad_SI_Label.Text + ", T3_ResidualBoltYieldPC = " + T3RBoltYield_Per_Label.Text 
                                                   + ", DetensioningBoltStress = " + DetenBoltStress_SI_Label.Text + ", DetensioningBoltLoad = " + DetenBoltLoad_SI_Label.Text + ",	"
                                                   + "   DetensioningBoltYieldPC = " + DetenBoltYield_Per_Label.Text + ", TensionPressure_FirstPass = " + Pass1AppliedPressure_SequenceTabValue_SI_Label.Text + ", TensionPressure_SecondPass = " + Pass2AppliedPressure_SequenceTabValue_SI_Label.Text + ", TensionPressure_ThirdPass = " + Pass3AppliedPressure_SequenceTabValue_SI_Label.Text + ", CheckingPass = " + CheckPass1AppliedPressure_SequenceTabValue_SI_Label.Text + ",   "
                                                   + "   Torque = " + TorqueValue + ", Coefficient_Friction = " + CoefficientValue_Label.Text + ", Bolt01 = " + "'1', Bolt02 = '" + Pass2Bolt_SequenceTabValue_Label.Text + "', Bolt03 = '" + Pass3Bolt_SequenceTabValue_Label.Text + "', Bolt04 = '" + Pass4Bolt_SequenceTabValue_Label.Text + "', "
                                                   + "   FirstPass_100 = " + FirstPass_100 + ", FirstPass_50 = " + FirstPass_50 + ", SecondPass_50 = " + SecondPass_50 + ", Max_Detensioning = " + Max_Detensioning + ", "
                                                   + "   ResidualBoltStress = " + ResidualBoltStress_TextBox.Text + ", Comments = '" + Commented + "', IsDetension = '" + IsDetensioned + "', CrossLoadingPC = " + CrossLoading_TextBox.Text + ", DetensioningPC = " + Detensioning_TextBox.Text
                                                   + "    WHERE ProjectReportId = " + ProjectId_Label.Text + " AND ApplicationId = " + AppId_Label.Text;

                        UpdateApplication_SQLiteCommand.ExecuteNonQuery();

                        try
                        {
                            UpdateApplication_SQLiteCommand.CommandText = "SELECT  PR.CustomerId, PR.ProjectReportId, ApplicationId, JointId, "
                                                                            + "   FlangeRating, Flange1_Abbreviation, Flange2_Abbreviation, "
                                                                            + "   (BoltThread || ' x ' ||	Pitch_TPI_Value) AS BoltSize, Number_of_Bolts, Material "
                                                                            + "   FROM ProjectReport   PR"
                                                                            + "   	INNER JOIN ProjectDetailedReport     PD 		ON		PD.ProjectReportId = PR.ProjectReportId "
                                                                            + "   	WHERE PR.ProjectReportId > 0  AND PR.ProjectReportId = " + ProjectId_Label.Text + ";";
                            DataSet AppGrid_DataSet = new DataSet();
                            SQLiteDataAdapter AppGrid_DataAdapter = new SQLiteDataAdapter(UpdateApplication_SQLiteCommand);


                            int AppGridResult = AppGrid_DataAdapter.Fill(AppGrid_DataSet);

                            // Data Grid Setup - Initially the Data Grid has not been binded to any data. Hence only first time you need to bind.
                            if (Project_DataGridView.DataSource != null)
                            {
                                Project_DataGridView.DataSource = null;
                            }

                            Project_DataGridView.AutoGenerateColumns = false;
                            Project_DataGridView.ColumnCount = 10;
                            Project_DataGridView.Columns[0].Visible = false;
                            Project_DataGridView.Columns[0].Name = "Customer_Id";
                            Project_DataGridView.Columns[0].HeaderText = "Customer Id";
                            Project_DataGridView.Columns[0].DataPropertyName = "CustomerId";
                            Project_DataGridView.Columns[1].Visible = false;
                            Project_DataGridView.Columns[1].Name = "ProjectReport_Id";
                            Project_DataGridView.Columns[1].HeaderText = "Project Report Id";
                            Project_DataGridView.Columns[1].DataPropertyName = "ProjectReportId";
                            Project_DataGridView.Columns[2].Visible = false;
                            Project_DataGridView.Columns[2].Name = "Application_Id";
                            Project_DataGridView.Columns[2].HeaderText = "Application Id";
                            Project_DataGridView.Columns[2].DataPropertyName = "ApplicationId";
                            Project_DataGridView.Columns[3].Visible = true;
                            Project_DataGridView.Columns[3].Name = "Joint_Id";
                            Project_DataGridView.Columns[3].HeaderText = "ID";
                            Project_DataGridView.Columns[3].DataPropertyName = "JointId";
                            Project_DataGridView.Columns[3].Width = 120;
                            Project_DataGridView.Columns[4].Visible = true;
                            Project_DataGridView.Columns[4].Name = "Flange_Rating";
                            Project_DataGridView.Columns[4].HeaderText = "Rating";
                            Project_DataGridView.Columns[4].DataPropertyName = "FlangeRating";
                            Project_DataGridView.Columns[4].Width = 180;
                            Project_DataGridView.Columns[5].Visible = true;
                            Project_DataGridView.Columns[5].Name = "Flange1";
                            Project_DataGridView.Columns[5].HeaderText = "Flange 1";
                            Project_DataGridView.Columns[5].DataPropertyName = "Flange1_Abbreviation";
                            Project_DataGridView.Columns[6].Visible = true;
                            Project_DataGridView.Columns[6].Name = "Flange2";
                            Project_DataGridView.Columns[6].HeaderText = "Flange 2";
                            Project_DataGridView.Columns[6].DataPropertyName = "Flange2_Abbreviation";
                            Project_DataGridView.Columns[7].Visible = true;
                            Project_DataGridView.Columns[7].Name = "Bolt_Diameter";
                            Project_DataGridView.Columns[7].HeaderText = "Bolt Size";
                            Project_DataGridView.Columns[7].DataPropertyName = "BoltSize";
                            Project_DataGridView.Columns[8].Visible = true;
                            Project_DataGridView.Columns[8].Name = "Bolts";
                            Project_DataGridView.Columns[8].HeaderText = "Bolts";
                            Project_DataGridView.Columns[8].DataPropertyName = "Number_of_Bolts";
                            Project_DataGridView.Columns[9].Visible = true;
                            Project_DataGridView.Columns[9].Name = "Bolt_Material";
                            Project_DataGridView.Columns[9].HeaderText = "Bolt Material";
                            Project_DataGridView.Columns[9].DataPropertyName = "Material";
                            Project_DataGridView.Columns[9].Width = 160;
                            Project_DataGridView.DataSource = AppGrid_DataSet.Tables[0];

                        }
                        catch (Exception AppRetrieval_Exception)
                        {
                            MessageBox.Show(AppRetrieval_Exception.Message + " while populating project data grid of new application.", "Error!");
                        }

                    }
                    catch (Exception AppUpdate_Exception)
                    {
                        MessageBox.Show(AppUpdate_Exception.Message + " while updating application.", "Fatal Error!!!");
                    }
                    finally
                    {
                        ConnectApplication.Close();
                        UpdateApplication_SQLiteCommand.Dispose();
                        ConnectApplication.Close();
                    }
                } 
                else
                {
                    MessageBox.Show("Please save the application by clicking 'Save As New Application' menuitem from 'Application' menu or 'Save Application as New' button on toolbar.", "Instruction!");
                }
            }
        }

        public void RemoveProject()
        {
            SQLiteConnection Connect_for_Delete = new SQLiteConnection("Data Source=" + BasePath + "\\WizBolt.db;New=False;");
            SQLiteCommand Delete_Command = new SQLiteCommand(Connect_for_Delete);
            
            try
            {
                Connect_for_Delete.Open();
                if ((ProjectId_Label.Text.Length > 0) && (Convert.ToInt32(ProjectId_Label.Text) > 0))
                {
                    DialogResult UserResponse = MessageBox.Show("This project will be parmanently deleted and you will never be able to get its contents again. This is irreversible process. Do you really want to delete?", "Warning!", MessageBoxButtons.YesNo);
                    if (UserResponse == DialogResult.Yes)
                    {
                            Delete_Command.CommandText = "DELETE FROM ProjectDetailedReport WHERE  ProjectReportId = " + ProjectId_Label.Text;
                            int DelResult = Delete_Command.ExecuteNonQuery();
                            if (DelResult >= 0)
                            {
                                AppId_Label.Text = "1";
                                Delete_Command.CommandText = "DELETE FROM ProjectReport WHERE  ProjectReportId = " + ProjectId_Label.Text;
                                int Outcome = Delete_Command.ExecuteNonQuery();
                                if (Outcome > 0)
                                {
                                    ProjectId_Label.Text = "0";
                                    CustomerId_Label.Text = "0";
                                    CustomerName_Label.Text = "Client Name";
                                    CustomerName_Label.Visible = false;
                                    LocationName_Label.Text = "Address";
                                    LocationName_Label.Visible = false;
                                   
                                    ProjectName_Label.Text = "Project Name";
                                    ProjectName_Label.Visible = false;
                                    Reference_Label.Text = "Project Reference";
                                    Reference_Label.Visible = false;
                                    ProjectDate_Label.Visible = false;
                                    EnteredProjectDate_label.Text = "20/07/2016";
                                    EnteredProjectDate_label.Visible = false;
                                    Project_DataGridView.DataSource = null;
                                    if (FileWithPath_Label.Text.Trim() != "File with path")
                                    {
                                        if (File.Exists(FileWithPath_Label.Text))
                                        {
                                            File.Delete(FileWithPath_Label.Text);
                                            FileWithPath_Label.Text = "File with path";
                                            MessageBox.Show("Project deleted successfully", "User Information");
                                        }
                                        else
                                        {
                                            MessageBox.Show("Project deleted successfully", "User Information");
                                        }
                                    }
                                    else
                                    {
                                        MessageBox.Show("Project deleted successfully", "User Information");
                                    }
                                    AssignInitially();
                                    JointId_TextBox.Text = "NEW APPLICATION";
                                }
                        }
                    }
                    
                }
                else
                {
                    MessageBox.Show("Currently, there isn't any project, either opened or created. Cannot remove any project.");
                }
            }
            catch (Exception ProjectDelete_Exception)
            {
                MessageBox.Show(ProjectDelete_Exception.Message + " while deleting project.", "Fatal Error!!!");
            }
            finally
            {
                Connect_for_Delete.Close();
                Delete_Command.Dispose();
                Connect_for_Delete.Dispose();
            }
        }

        public void RemoveApplication()
        {
            SQLiteConnection Connect_for_App = new SQLiteConnection("Data Source=" + BasePath + "\\WizBolt.db;New=False;");
            SQLiteCommand RemoveApp_Command = new SQLiteCommand(Connect_for_App);
            Connect_for_App.Open();

            try
            {
                if ((ProjectId_Label.Text.Length > 0) && (Convert.ToInt32(ProjectId_Label.Text.Trim()) > 0))
                {
                    if ((AppId_Label.Text.Length > 0) && (Convert.ToInt32(AppId_Label.Text.Trim()) > 0))
                    {
                        RemoveApp_Command.CommandText = " DELETE FROM ProjectDetailedReport WHERE  ProjectReportId = " + ProjectId_Label.Text.Trim()
                                                            + " AND ApplicationId = " + AppId_Label.Text.Trim();
                        int AppResult = RemoveApp_Command.ExecuteNonQuery();
                        if (AppResult > 0)
                        {
                            RemoveApp_Command.CommandText = "SELECT ApplicationId FROM ProjectDetailedReport WHERE ProjectReportId  = " + ProjectId_Label.Text.Trim();
                            SQLiteDataReader App_DataReader = RemoveApp_Command.ExecuteReader();
                                if (App_DataReader.Read())
                                {
                                    if (App_DataReader.IsDBNull(0))
                                    {
                                       Project_DataGridView.DataSource = null;
                                       AppId_Label.Text = "1";
                                       IsInitial = true;
                                       if (!App_DataReader.IsClosed)
                                       {
                                           App_DataReader.Close();
                                           App_DataReader.Dispose();
                                       }
                                       AssignInitially();
                                       JointId_TextBox.Text = "NEW APPLICATION";
                                       
                                    }
                                    else
                                    {
                                        AppId_Label.Text = (App_DataReader.GetInt32(0)).ToString();
                                        if (!App_DataReader.IsClosed)
                                        {
                                            App_DataReader.Close();
                                            App_DataReader.Dispose();
                                        }
                                         RemoveApp_Command.CommandText = "SELECT  PR.CustomerId, PR.ProjectReportId, ApplicationId, JointId, "
                                                                    + "   FlangeRating, Flange1_Abbreviation, Flange2_Abbreviation, "
                                                                    + "   (BoltThread || ' x ' ||	Pitch_TPI_Value) AS BoltSize, Number_of_Bolts, Material "
                                                                    + "   FROM ProjectReport   PR"
                                                                    + "   	INNER JOIN ProjectDetailedReport     PD 		ON		PD.ProjectReportId = PR.ProjectReportId "
                                                                    + "   	WHERE PR.ProjectReportId > 0  AND PR.ProjectReportId = " + ProjectId_Label.Text + ";";
                            DataSet AppGrid_DataSet = new DataSet();
                            SQLiteDataAdapter AppGrid_DataAdapter = new SQLiteDataAdapter(RemoveApp_Command);

                            // Setting DataGridView
                            int AppGridResult = AppGrid_DataAdapter.Fill(AppGrid_DataSet);
                            if (AppGridResult > 0)
                            {
                                if (Project_DataGridView.DataSource != null)
                                {
                                    Project_DataGridView.DataSource = null;
                                }
                                Project_DataGridView.AutoGenerateColumns = false;
                                Project_DataGridView.ColumnCount = 10;
                                Project_DataGridView.Columns[0].Visible = false;
                                Project_DataGridView.Columns[0].Name = "Customer_Id";
                                Project_DataGridView.Columns[0].HeaderText = "Customer Id";
                                Project_DataGridView.Columns[0].DataPropertyName = "CustomerId";
                                Project_DataGridView.Columns[1].Visible = false;
                                Project_DataGridView.Columns[1].Name = "ProjectReport_Id";
                                Project_DataGridView.Columns[1].HeaderText = "Project Report Id";
                                Project_DataGridView.Columns[1].DataPropertyName = "ProjectReportId";
                                Project_DataGridView.Columns[2].Visible = false;
                                Project_DataGridView.Columns[2].Name = "Application_Id";
                                Project_DataGridView.Columns[2].HeaderText = "Application Id";
                                Project_DataGridView.Columns[2].DataPropertyName = "ApplicationId";
                                Project_DataGridView.Columns[3].Visible = true;
                                Project_DataGridView.Columns[3].Name = "Joint_Id";
                                Project_DataGridView.Columns[3].HeaderText = "ID";
                                Project_DataGridView.Columns[3].DataPropertyName = "JointId";
                                Project_DataGridView.Columns[3].Width = 120;
                                Project_DataGridView.Columns[4].Visible = true;
                                Project_DataGridView.Columns[4].Name = "Flange_Rating";
                                Project_DataGridView.Columns[4].HeaderText = "Rating";
                                Project_DataGridView.Columns[4].DataPropertyName = "FlangeRating";
                                Project_DataGridView.Columns[4].Width = 180;
                                Project_DataGridView.Columns[5].Visible = true;
                                Project_DataGridView.Columns[5].Name = "Flange1";
                                Project_DataGridView.Columns[5].HeaderText = "Flange 1";
                                Project_DataGridView.Columns[5].DataPropertyName = "Flange1_Abbreviation";
                                Project_DataGridView.Columns[6].Visible = true;
                                Project_DataGridView.Columns[6].Name = "Flange2";
                                Project_DataGridView.Columns[6].HeaderText = "Flange 2";
                                Project_DataGridView.Columns[6].DataPropertyName = "Flange2_Abbreviation";
                                Project_DataGridView.Columns[7].Visible = true;
                                Project_DataGridView.Columns[7].Name = "Bolt_Diameter";
                                Project_DataGridView.Columns[7].HeaderText = "Bolt Size";
                                Project_DataGridView.Columns[7].DataPropertyName = "BoltSize";
                                Project_DataGridView.Columns[8].Visible = true;
                                Project_DataGridView.Columns[8].Name = "Bolts";
                                Project_DataGridView.Columns[8].HeaderText = "Bolts";
                                Project_DataGridView.Columns[8].DataPropertyName = "Number_of_Bolts";
                                Project_DataGridView.Columns[9].Visible = true;
                                Project_DataGridView.Columns[9].Name = "Bolt_Material";
                                Project_DataGridView.Columns[9].HeaderText = "Bolt Material";
                                Project_DataGridView.Columns[9].DataPropertyName = "Material";
                                Project_DataGridView.Columns[9].Width = 160;
                                Project_DataGridView.DataSource = AppGrid_DataSet.Tables[0];

                                // Set first application in display
                                SetApplication();
                            }
                          }
                         }
                         else
                         {
                             Project_DataGridView.DataSource = null;
                         }
                        }
                    }
                }
            }
            catch (Exception RemoveApp_Exception)
            {
                MessageBox.Show(RemoveApp_Exception.Message + " while deleting application.", "Fatal Error!!!");
            }
            finally
            {
                Connect_for_App.Close();
                RemoveApp_Command.Dispose();
                Connect_for_App.Dispose();
            }
        }

        public void SaveClose_Project()
        {
            SaveProject();
            CloseProject();
        }

        public void CloseProject()
        {
            SQLiteConnection Connect_for_Delete = new SQLiteConnection("Data Source=" + BasePath + "\\WizBolt.db;New=False;");
            SQLiteCommand Delete_Command = new SQLiteCommand(Connect_for_Delete);
            Connect_for_Delete.Open();
            Delete_Command.CommandText = "DELETE FROM ProjectDetailedReport WHERE  ProjectReportId > 0;";
            int DelResult = Delete_Command.ExecuteNonQuery();
            if (DelResult >= 0)
            {
                Delete_Command.CommandText = "DELETE FROM ProjectReport WHERE  ProjectReportId > 0;";
                int Outcome = Delete_Command.ExecuteNonQuery();
                if (Outcome > 0)
                {
                    Delete_Command.CommandText = "DELETE FROM Customer;";
                    int CustomerResult = Delete_Command.ExecuteNonQuery();
                    Delete_Command.CommandText = "DELETE FROM Projects;";
                    int ProjectResult = Delete_Command.ExecuteNonQuery();
                }
            }
            Connect_for_Delete.Close();

            ProjectId_Label.Text = "0";
            AppId_Label.Text = "1";
            CustomerId_Label.Text = "0";
            CustomerName_Label.Text = "Client Name";
            CustomerName_Label.Visible = false;
            LocationName_Label.Text = "Address";
            LocationName_Label.Visible = false;
            ProjectName_Label.Text = "Project Name";
            ProjectName_Label.Visible = false;
            Reference_Label.Text = "Project Reference";
            Reference_Label.Visible = false;
            ProjectDate_Label.Visible = false;
            EnteredProjectDate_label.Text = "20/07/2016";
            EnteredProjectDate_label.Visible = false;
            ProjectEndDate_Label.Text = "Not Entered";
            ProjectEndDate_Label.Visible = false;
            ProjectEndDateTitle_Label.Visible = false;
            FileWithPath_Label.Text = "File with path";
            Project_DataGridView.DataSource = null;
            DefaultStress_Button.Visible = false;
            //IsInitial = true;
            AssignInitially();
            JointId_TextBox.Text = "NEW APPLICATION";

        }

        private void Pitch_TextBox_TextChanged(object sender, EventArgs e)
        {
            if (!IsInitial)
            {
                decimal StressArea = 0M;
                decimal Torque = 0M;
                decimal Bolt_Length = 0M;
                decimal BoltDiameter_for_StressArea = 0M;
                decimal ClampLength = 0M;
                decimal TopClamp = 0M;
                decimal BottomClamp = 0M;
                decimal GasketGap = 0M;
                decimal SpacerThickness = 0M;
                decimal pitch_equivalent = 0M;
                decimal LTFValue = 0M;
                decimal ToolPressureArea = 0M;
                decimal ResidualBoltStress = 0M;
                decimal Coefficient_Friction = 0M;
                string TPA = string.Empty;
                int SelectedToolIndex = 0;

                //if (!IsSetting)
                //{
                    decimal Pitch = 0M;
                    string BoltDiameter = string.Empty;
                    string Bolt_Diameter = string.Empty;
                    if (!IsManual)
                    {
                        if ((Pitch_TextBox.Text.Trim() != "") && (Pitch_TextBox.Text.Trim().Length > 0))
                        {
                            Pitch = Convert.ToDecimal(Pitch_TextBox.Text.Trim());
                           
                        }
                        if (Pitch > 0)
                        {
                            pitch_equivalent = 1.0M / Pitch;
                        }
                    
                        
                        if (Clamp1_TextBox.Text.Length > 0)
                        {
                            TopClamp = Convert.ToDecimal(Clamp1_TextBox.Text);
                        }
                        if (Clamp2_TextBox.Text.Length > 0)
                        {
                            BottomClamp = Convert.ToDecimal(Clamp2_TextBox.Text);
                        }
                        if (GasketGap_TextBox.Text.Length > 0)
                        {
                            GasketGap = Convert.ToDecimal(GasketGap_TextBox.Text);
                        }
                        if (CustomSpacer_TextBox.Text.Length > 0)
                        {
                            SpacerThickness = Convert.ToDecimal(CustomSpacer_TextBox.Text);
                        }
                        if ((ResidualBoltStress_TextBox.Text.Trim() != "") && (ResidualBoltStress_TextBox.Text.Trim().Length > 0))
                        {
                            ResidualBoltStress = Convert.ToDecimal(ResidualBoltStress_TextBox.Text.Trim());
                        }
                        if ((Coeff_TextBox.Text.Trim() != "") && (Coeff_TextBox.Text.Trim().Length > 0))
                        {
                            Coefficient_Friction = Convert.ToDecimal(Coeff_TextBox.Text.Trim());
                        }
                        ClampLength = TopClamp + BottomClamp + GasketGap + SpacerThickness;

                        if ((Pitch_TextBox.Text.Trim() != "") && (Pitch_TextBox.Text.Trim().Length > 0))
                        {
                            Pitch = Convert.ToDecimal(Pitch_TextBox.Text.Trim());
                           
                        }
                        if (Pitch > 0)
                        {
                            pitch_equivalent = 1.0M / Pitch;
                        }
                        LTFValue = Calculate_LTF();

                        // Info Panel
                        LTF_Label.Text = LTFValue.ToString();
                        PitchInfo_Label.Text = Pitch_TextBox.Text.Trim();
                        NoOfBoltsInfo_Label.Text = NumberOfBolts_TextBox.Text;

                        try
                        {
                            Bolt_Diameter = ((DataRowView)BoltThread_ComboBox.SelectedItem)["BoltThread"].ToString();
                        }
                        catch (Exception BoltDiameter_Exception)
                        {
                            MessageBox.Show(BoltDiameter_Exception.Message + " in pitch change event.");
                        }

                        // Bolt Tab
                        NominalBoltDiameterValueTab_Label.Text = Bolt_Diameter;
                        NumberOfBoltsValueTab_Label.Text = NumberOfBolts_TextBox.Text.Trim();

                        if (Bolt_Diameter.Substring(0, 1) == "M")
                        {
                            Pitch_Label.Text = "Pitch";

                            // Info Panel
                            NominalThreadSizeInfo_Label.Text = Bolt_Diameter + "x" + Pitch_TextBox.Text.Trim();
                            TPI_Label.Text = Pitch_Label.Text;
                            
                            BoltDiameter_for_StressArea = Convert.ToDecimal(Bolt_Diameter.Substring(1));
                            StressArea = 0.7854M * ((BoltDiameter_for_StressArea - 0.938194M * Pitch) * (BoltDiameter_for_StressArea - 0.938194M * Pitch));
                            TensileStressArea_SI_Label.Text = Math.Round(StressArea, 2, MidpointRounding.AwayFromZero).ToString() + " Sq mm";
                            TensileStressArea_In_Label.Text = Math.Round(Convert.ToDecimal((StressArea / 645.16M)), 4, MidpointRounding.AwayFromZero).ToString() + " Sq in";
                            
                            // Bolt Tab
                            TensileStressAreaValueTab_In_Label.Text = TensileStressArea_In_Label.Text;
                            TensileStressAreaValueTab_SI_Label.Text = TensileStressArea_SI_Label.Text;
                            if (UnitSystem_Button.Text == "SAE")
                            {
                                // Info Panel
                                ClampLengthInfo_Label.Text = Math.Round(ClampLength, 4, MidpointRounding.AwayFromZero).ToString() + " in";
                                ClampLengthInfo_mm_Label.Text = Math.Round((ClampLength * 25.4M), 2, MidpointRounding.AwayFromZero).ToString() + " mm";
                                Bolt_Length = Math.Round((ClampLength + 3 * (BoltDiameter_for_StressArea / 25.4M) + 3 * (Pitch / 25.4M)), 4, MidpointRounding.AwayFromZero);
                                BoltLength_In_Label.Text = Bolt_Length.ToString() + " in";
                                BoltLength_SI_Label.Text = Math.Round((Bolt_Length * 25.4M), 2, MidpointRounding.AwayFromZero).ToString() + " mm";
                                // Bolt Tab
                                BoltLengthValueTab_In_Label.Text = BoltLength_In_Label.Text;
                                BoltLengthValueTab_SI_Label.Text = BoltLength_SI_Label.Text;
                                // Torque = (ResidualBoltStress/145.037738007M) * Coefficient_Friction * BoltDiameter_for_StressArea;
                                Torque = (ResidualBoltStress / 145.037738007M) * ((Pitch / (2M * 3.14159M)) + (Coefficient_Friction * BoltDiameter_for_StressArea) / 1.7320508M + (Coefficient_Friction * BoltDiameter_for_StressArea * 1.5M / 2M ));
                                // Info Panel
                                Torque_In_Label.Text = Math.Round((Torque / 1.35581794833M), 0, MidpointRounding.AwayFromZero).ToString() + " lb-ft";
                                Torque_SI_Label.Text = Math.Round(Torque, 0, MidpointRounding.AwayFromZero).ToString() + " N-m";
                                // Torque Tab
                                TorqueTab_SI_Label.Text = Torque_SI_Label.Text;
                                TorqueValueTab_In_Label.Text = Torque_In_Label.Text;
                                
                                // Sequence Tab
                                Torque_SI_Sequence_Label.Text = Torque_SI_Label.Text;
                                Torque_In_Sequence_Label.Text = Torque_In_Label.Text;
                            }
                            else
                            {
                                // Info Panel
                                ClampLengthInfo_Label.Text = (Math.Round((ClampLength / 25.4M), 4, MidpointRounding.AwayFromZero)).ToString() + " in";
                                ClampLengthInfo_mm_Label.Text = Math.Round(ClampLength, 2, MidpointRounding.AwayFromZero).ToString() + " mm";
                                Bolt_Length = ClampLength + 3 * BoltDiameter_for_StressArea + 3 * Pitch;
                                BoltLength_SI_Label.Text = Math.Round(Bolt_Length, 2, MidpointRounding.AwayFromZero).ToString() + " mm";
                                BoltLength_In_Label.Text = Math.Round((Bolt_Length / 25.4M), 4, MidpointRounding.AwayFromZero).ToString() + " in";
                                // Bolt Tab
                                BoltLengthValueTab_In_Label.Text = BoltLength_In_Label.Text;
                                BoltLengthValueTab_SI_Label.Text = BoltLength_SI_Label.Text;

                               // Torque = (ResidualBoltStress * Coefficient_Friction * BoltDiameter_for_StressArea);
                                Torque = (ResidualBoltStress) * ((Pitch / (2M * 3.14159M)) + (Coefficient_Friction * BoltDiameter_for_StressArea) / 1.7320508M + (Coefficient_Friction * BoltDiameter_for_StressArea * 1.5M / 2M));
                                // Info Panel
                                Torque_In_Label.Text = Math.Round((Torque * 0.737562149277M), 0, MidpointRounding.AwayFromZero).ToString() + " lb-ft";
                                Torque_SI_Label.Text = Math.Round(Torque, 0, MidpointRounding.AwayFromZero).ToString() + " N-m";
                                // Torque Tab
                                TorqueTab_SI_Label.Text = Math.Round(Torque, 0, MidpointRounding.AwayFromZero).ToString() + " N-m";
                                TorqueValueTab_In_Label.Text = Math.Round((Torque * 0.737562149277M), 0, MidpointRounding.AwayFromZero).ToString() + " lb-ft";

                                // Sequence Tab
                                Torque_SI_Sequence_Label.Text = TorqueTab_SI_Label.Text;
                                Torque_In_Sequence_Label.Text = TorqueValueTab_In_Label.Text;
                            }
                        }
                        else
                        {
                            Pitch_Label.Text = "T. P. I.";
                            //Info Panel
                            NominalThreadSizeInfo_Label.Text = Bolt_Diameter + "-" + Pitch + "UN";
                            TPI_Label.Text = Pitch_Label.Text;
                           
                            Bolt_Diameter = Bolt_Diameter.Substring(0, (Bolt_Diameter.Length - 1)).Trim();
                            BoltDiameter_for_StressArea = Convert_Decimal(Bolt_Diameter);
                            BoltDiameter_for_StressArea = Math.Round((BoltDiameter_for_StressArea * 25.4M), 4, MidpointRounding.AwayFromZero);
                            if (Pitch > 0)
                            {
                                StressArea = 0.7854M * (BoltDiameter_for_StressArea - (0.9743M / Pitch)) * (BoltDiameter_for_StressArea - (0.9743M / Pitch));
                            }
                            TensileStressArea_SI_Label.Text = Math.Round(StressArea, 2, MidpointRounding.AwayFromZero).ToString() + " Sq mm";
                            TensileStressArea_In_Label.Text = Math.Round(Convert.ToDecimal((StressArea / 645.16M)), 4, MidpointRounding.AwayFromZero).ToString() + " Sq in";
                            
                            // Bolt Tab
                            TensileStressAreaValueTab_In_Label.Text = TensileStressArea_In_Label.Text;
                            TensileStressAreaValueTab_SI_Label.Text = TensileStressArea_SI_Label.Text;

                            if (UnitSystem_Button.Text == "SAE")
                            {
                                // Info Panel
                                ClampLengthInfo_Label.Text = Math.Round(ClampLength, 4, MidpointRounding.AwayFromZero).ToString() + " in";
                                ClampLengthInfo_mm_Label.Text = Math.Round((ClampLength * 25.4M), 2, MidpointRounding.AwayFromZero).ToString() + " mm";

                                Bolt_Length = ClampLength + 3 * (BoltDiameter_for_StressArea / 25.4M) + 3 * pitch_equivalent;

                                BoltLength_In_Label.Text = Math.Round(Bolt_Length, 4, MidpointRounding.AwayFromZero).ToString() + " in";
                                BoltLength_SI_Label.Text = Math.Round((Bolt_Length * 25.4M), 2, MidpointRounding.AwayFromZero).ToString() + " mm";
                                // Bolt Tab
                                BoltLengthValueTab_In_Label.Text = BoltLength_In_Label.Text;
                                BoltLengthValueTab_SI_Label.Text = BoltLength_SI_Label.Text;

                               // Torque = ((ResidualBoltStress / 145.037738007M) * Coefficient_Friction * BoltDiameter_for_StressArea);
                                Torque = (ResidualBoltStress / 145.037738007M) * ((pitch_equivalent * 25.4M / (2M * 3.14159M)) + (Coefficient_Friction * BoltDiameter_for_StressArea) / 1.7320508M + (Coefficient_Friction * BoltDiameter_for_StressArea * 1.5M / 2M));

                                // Info Panel
                                Torque_In_Label.Text = Math.Round((Torque / 1.35581794833M), 0, MidpointRounding.AwayFromZero).ToString() + " lb-ft";
                                Torque_SI_Label.Text = Math.Round(Torque, 0, MidpointRounding.AwayFromZero).ToString() + " N-m";
                                // Torque Tab
                                TorqueTab_SI_Label.Text = Torque_SI_Label.Text;
                                TorqueValueTab_In_Label.Text = Torque_In_Label.Text;

                                // Sequence Tab
                                Torque_SI_Sequence_Label.Text = Torque_SI_Label.Text;
                                Torque_In_Sequence_Label.Text = Torque_In_Label.Text;
                            }
                            else
                            {
                                // Info Panel
                                ClampLengthInfo_Label.Text = (Math.Round((ClampLength / 25.4M), 4, MidpointRounding.AwayFromZero)).ToString() + " in";
                                ClampLengthInfo_mm_Label.Text = Math.Round(ClampLength, 2, MidpointRounding.AwayFromZero).ToString() + " mm";
                                
                                Bolt_Length = ClampLength + 3 * BoltDiameter_for_StressArea + 3 * pitch_equivalent * 25.4M;

                                BoltLength_SI_Label.Text = Math.Round(Bolt_Length, 2, MidpointRounding.AwayFromZero).ToString() + " mm";
                                BoltLength_In_Label.Text = Math.Round((Bolt_Length / 25.4M), 4, MidpointRounding.AwayFromZero).ToString() + " in";
                                // Bolt Tab
                                BoltLengthValueTab_In_Label.Text = BoltLength_In_Label.Text;
                                BoltLengthValueTab_SI_Label.Text = BoltLength_SI_Label.Text;

                                //Torque = (ResidualBoltStress * Coefficient_Friction * BoltDiameter_for_StressArea);
                                Torque = (ResidualBoltStress) * ((pitch_equivalent * 25.4M / (2M * 3.14159M)) + (Coefficient_Friction * BoltDiameter_for_StressArea) / 1.7320508M + (Coefficient_Friction * BoltDiameter_for_StressArea * 1.5M / 2M));
                                // Info Panel
                                Torque_In_Label.Text = Math.Round((Torque * 0.737562149277M), 0, MidpointRounding.AwayFromZero).ToString() + " lb-ft";
                                Torque_SI_Label.Text = Math.Round(Torque, 0, MidpointRounding.AwayFromZero).ToString() + " N-m";
                                // Torque Tab
                                TorqueTab_SI_Label.Text = Torque_SI_Label.Text;
                                TorqueValueTab_In_Label.Text = Torque_In_Label.Text;

                                // Sequence Tab
                                Torque_SI_Sequence_Label.Text = Torque_SI_Label.Text;
                                Torque_In_Sequence_Label.Text = Torque_In_Label.Text;
                            }
                        }

                        /*
                        * Tensioner Tool Information: When tensioner tool is available then this info. Also remember TensionerTool_SelectedIndexChanged event 
                        * is activated only when user changes tool from tool dropdown, else it is prevented by making IsBoltThread = true as it is initiated from
                        * BoltThread dropdown.
                        */
                        // Tensioner Tool
                        if (TensionerTool_ComboBox.DataSource != null)
                        {
                            DataTable Tool_Table = (DataTable)TensionerTool_ComboBox.DataSource;
                            SelectedToolIndex = TensionerTool_ComboBox.SelectedIndex;
                            // Info Panel
                            ToolId_Info_Label.Text = ((DataRowView)TensionerTool_ComboBox.SelectedItem)["ModelNumber"].ToString();
                            TensionerBolt_Info_Label.Text = NominalThreadSizeInfo_Label.Text;
                            TPA = Tool_Table.Rows[SelectedToolIndex]["HydraulicArea"].ToString();
                            ToolPressureArea = Math.Round((Convert.ToDecimal(TPA) / 645.16M), 4, MidpointRounding.AwayFromZero);
                            ToolPressureArea_Info_Label.Text = ToolPressureArea.ToString() + " Sq in";
                            ToolPressureArea_mm_Label.Text = TPA + " Sq mm";
                         }
                         else
                         {
                           // Info Panel
                           ToolId_Info_Label.Text = string.Empty;
                           TensionerBolt_Info_Label.Text = string.Empty;
                           ToolPressureArea_Info_Label.Text = string.Empty;
                           ToolPressureArea_mm_Label.Text = string.Empty;
                         }
                        // Tool Pressures
                        OperateTool();
                       // Graph
                      DrawGraph();
                    }
                    else    // For Manual input
                    {
                        // Info Panel
                        PitchInfo_Label.Text = Pitch_TextBox.Text.Trim();
                        NoOfBoltsInfo_Label.Text = NumberOfBolts_TextBox.Text;

                        try
                        {
                            Bolt_Diameter = ((DataRowView)BoltThread_ComboBox.SelectedItem)["BoltThread"].ToString();
                        }
                        catch (Exception BoltDiameter_Exception)
                        {
                            MessageBox.Show(BoltDiameter_Exception.Message + " in pitch change event.");
                        }

                        // Bolt Tab
                        NominalBoltDiameterValueTab_Label.Text = Bolt_Diameter;
                        NumberOfBoltsValueTab_Label.Text = NumberOfBolts_TextBox.Text.Trim();

                        if (Bolt_Diameter.Substring(0, 1) == "M")
                        {
                            Pitch_Label.Text = "Pitch";

                            // Info Panel
                            NominalThreadSizeInfo_Label.Text = Bolt_Diameter + "x" + Pitch_TextBox.Text.Trim();
                            TPI_Label.Text = Pitch_Label.Text;

                            BoltDiameter_for_StressArea = Convert.ToDecimal(Bolt_Diameter.Substring(1));
                            StressArea = 0.7854M * (BoltDiameter_for_StressArea - 0.938194M * Pitch) * (BoltDiameter_for_StressArea - 0.938194M * Pitch);
                            TensileStressArea_SI_Label.Text = Math.Round(StressArea, 2, MidpointRounding.AwayFromZero).ToString() + " Sq mm";
                            TensileStressArea_In_Label.Text = Math.Round(Convert.ToDecimal((StressArea / 645.16M)), 4, MidpointRounding.AwayFromZero).ToString() + " Sq in";

                            // Bolt Tab
                            TensileStressAreaValueTab_In_Label.Text = TensileStressArea_In_Label.Text;
                            TensileStressAreaValueTab_SI_Label.Text = TensileStressArea_SI_Label.Text;
                        }
                        else
                        {
                            Pitch_Label.Text = "T. P. I.";
                            //Info Panel
                            NominalThreadSizeInfo_Label.Text = Bolt_Diameter + "-" + Pitch + "UN";
                            TPI_Label.Text = Pitch_Label.Text;

                            Bolt_Diameter = Bolt_Diameter.Substring(0, (Bolt_Diameter.Length - 1)).Trim();
                            BoltDiameter_for_StressArea = Convert_Decimal(Bolt_Diameter);
                            BoltDiameter_for_StressArea = Math.Round((BoltDiameter_for_StressArea * 25.4M), 4, MidpointRounding.AwayFromZero);
                            if (Pitch > 0)
                            {
                                StressArea = 0.7854M * (BoltDiameter_for_StressArea - (0.9743M / Pitch)) * (BoltDiameter_for_StressArea - (0.9743M / Pitch));
                            }
                            TensileStressArea_SI_Label.Text = Math.Round(StressArea, 2, MidpointRounding.AwayFromZero).ToString() + " Sq mm";
                            TensileStressArea_In_Label.Text = Math.Round(Convert.ToDecimal((StressArea / 645.16M)), 4, MidpointRounding.AwayFromZero).ToString() + " Sq in";

                            // Bolt Tab
                            TensileStressAreaValueTab_In_Label.Text = TensileStressArea_In_Label.Text;
                            TensileStressAreaValueTab_SI_Label.Text = TensileStressArea_SI_Label.Text;
                        }
                        TensionerBolt_Info_Label.Text = NominalThreadSizeInfo_Label.Text;
                    }
                //}
            }
        }

        private void TensileStressArea_RadioButton_CheckedChanged(object sender, EventArgs e)
        {
            if (!IsManual)
            {
                decimal LTFValue = 1.11M;
                decimal NominalDiameter = 0M;
                decimal Pitch = 0M;
                decimal Pitch_Equivalent = 0M;
                decimal MinorDiameter = 0M;
                decimal MinorDiameterArea = 0M;
                decimal TensileStressArea = 0M;
                decimal Coef_Friction = 0M;
                decimal ResidualStress = 0M;
                decimal Torque = 0M;
                decimal ToolPressureArea = 0M;
                decimal BoltYield = 0M;
                decimal CrossLoad = 0M;
                string Bolt_Diameter = string.Empty;

                if ((LTF_Label.Text.Trim() != "") && (LTF_Label.Text.Trim().Length > 0))
                {
                    LTFValue = Convert.ToDecimal(LTF_Label.Text.Trim());
                }
                if ((CoefficientValue_Label.Text.Trim() != "") && (CoefficientValue_Label.Text.Trim().Length > 0))
                {
                    Coef_Friction = Convert.ToDecimal(CoefficientValue_Label.Text.Trim());
                }
                if ((T3RBoltStress_SI_Label.Text.Trim() != "") && (T3RBoltStress_SI_Label.Text.Trim().Length > 0))
                {
                    ResidualStress = Convert.ToDecimal(T3RBoltStress_SI_Label.Text.Trim());
                }
                if ((ToolPressureArea_mm_Label.Text.Trim() != "") && (ToolPressureArea_mm_Label.Text.Trim().Length > 6))
                {
                    ToolPressureArea = Convert.ToDecimal(ToolPressureArea_mm_Label.Text.Substring(0, ToolPressureArea_mm_Label.Text.Length - 6));
                }
                if ((BoltYield_SI_Label.Text.Trim() != "") && (BoltYield_SI_Label.Text.Trim().Length > 3))
                {
                    BoltYield = Convert.ToDecimal(BoltYield_SI_Label.Text.Substring(0, BoltYield_SI_Label.Text.Length - 3));
                }
                if ((CrossLoading_TextBox.Text.Trim() != "") && (CrossLoading_TextBox.Text.Trim().Length > 0))
                {
                    CrossLoad = Convert.ToDecimal(CrossLoading_TextBox.Text);
                }
                Bolt_Diameter = ((DataRowView)BoltThread_ComboBox.SelectedItem)["BoltThread"].ToString();

                if (TensileStressArea_RadioButton.Checked)
                {
                    MinorDiaArea_RadioButton.Checked = false;
                    if (Bolt_Diameter.Substring(0, 1) == "M")
                    {
                        // Units are in mm
                        NominalDiameter = Convert.ToDecimal(Bolt_Diameter.Substring(1));
                        if ((Pitch_TextBox.Text.Trim() != "") && (Pitch_TextBox.Text.Trim().Length > 0))
                        {
                            Pitch = Convert.ToDecimal(Pitch_TextBox.Text.Trim());
                            TensileStressArea = 0.7854M * (NominalDiameter - (0.938194M * Pitch)) * (NominalDiameter - (0.938194M * Pitch));
                            TensileStressArea_Title_Label.Text = "Tensile Stress Area";
                            TensileStressArea_In_Label.Text = Math.Round((TensileStressArea / 645.16M), 4, MidpointRounding.AwayFromZero).ToString() + " Sq in";
                            TensileStressArea_SI_Label.Text = Math.Round(TensileStressArea, 2, MidpointRounding.AwayFromZero).ToString() + " Sq mm";
                            TensileStressAreaTitleTab_Label.Text = "Tensile Stress Area";
                            TensileStressAreaValueTab_In_Label.Text = TensileStressArea_In_Label.Text;
                            TensileStressAreaValueTab_SI_Label.Text = TensileStressArea_SI_Label.Text;
                            Torque = Math.Round((ResidualStress * NominalDiameter * Coef_Friction), 0, MidpointRounding.AwayFromZero);
                            Torque_SI_Label.Text = Torque.ToString() + " N-m";
                            TorqueTab_SI_Label.Text = Torque_SI_Label.Text;
                            Torque_In_Label.Text = Math.Round((Torque / 1.35581794833M), 0, MidpointRounding.AwayFromZero).ToString() + " ft-lbs";
                            TorqueValueTab_In_Label.Text = Torque_In_Label.Text;
                            Torque_SI_Sequence_Label.Text = Torque_SI_Label.Text;
                            Torque_In_Sequence_Label.Text = Torque_In_Label.Text;
                        }
                        else
                        {
                            MessageBox.Show("Pitch (No Threads Per Inch) value is wrong. Only non-zero numeral allowed", "Error!");
                        }
                    }
                    else
                    {
                        // Units are in inch
                        Bolt_Diameter = Bolt_Diameter.Substring(0, (Bolt_Diameter.Length - 1)).Trim();
                        NominalDiameter = Convert_Decimal(Bolt_Diameter);
                        if ((Pitch_TextBox.Text.Trim() != "") && (Pitch_TextBox.Text.Trim().Length > 0))
                        {
                            Pitch = Convert.ToDecimal(Pitch_TextBox.Text.Trim());
                            TensileStressArea = 0.7854M * (NominalDiameter - (0.9743M / Pitch)) * (NominalDiameter - (0.9743M / Pitch));
                            TensileStressArea_Title_Label.Text = "Tensile Stress Area";
                            TensileStressArea_In_Label.Text = Math.Round(TensileStressArea, 4, MidpointRounding.AwayFromZero).ToString() + " Sq in";
                            TensileStressArea_SI_Label.Text = Math.Round((TensileStressArea * 645.16M), 2, MidpointRounding.AwayFromZero).ToString() + " Sq mm";
                            TensileStressAreaTitleTab_Label.Text = "Tensile Stress Area";
                            TensileStressAreaValueTab_In_Label.Text = TensileStressArea_In_Label.Text;
                            TensileStressAreaValueTab_SI_Label.Text = TensileStressArea_SI_Label.Text;
                            Torque = Math.Round((ResidualStress * (NominalDiameter * 25.4M) * Coef_Friction), 0, MidpointRounding.AwayFromZero);
                            Torque_SI_Label.Text = Torque.ToString() + " N-m";
                            TorqueTab_SI_Label.Text = Torque_SI_Label.Text;
                            Torque_In_Label.Text = Math.Round((Torque / 1.35581794833M), 0, MidpointRounding.AwayFromZero).ToString() + " ft-lbs";
                            TorqueValueTab_In_Label.Text = Torque_In_Label.Text;
                            TensileStressArea = TensileStressArea * 645.16M;
                            Torque_SI_Sequence_Label.Text = Torque_SI_Label.Text;
                            Torque_In_Sequence_Label.Text = Torque_In_Label.Text;
                        }
                        else
                        {
                            MessageBox.Show("Threads Per Inch value is wrong. Only non-zero numeral allowed. Do not enter pitch.", "Error!");
                        }
                    }
                    OperateTool();
                }
                else
                {
                    TensileStressArea_RadioButton.Checked = false;
                    if (Bolt_Diameter.Substring(0, 1) == "M")
                    {
                        // All units are in mm here.
                        NominalDiameter = Convert.ToDecimal(Bolt_Diameter.Substring(1));
                        if ((Pitch_TextBox.Text.Trim().Length > 0) && (Convert.ToDecimal(Pitch_TextBox.Text.Trim()) > 0M))
                        {
                            Pitch = Convert.ToDecimal(Pitch_TextBox.Text.Trim());
                            MinorDiameter = NominalDiameter - ((10M / 16M) * Convert.ToDecimal(Math.Sqrt(3)) * Pitch);
                            MinorDiameterArea = ((22M / 7M) * Convert.ToDecimal(Math.Pow(Convert.ToDouble(MinorDiameter) / 2, 2)));
                            GlobalMinorDiameterArea = MinorDiameterArea;
                            TensileStressArea_Title_Label.Text = "Minor Diameter Area";
                            TensileStressArea_In_Label.Text = Math.Round((MinorDiameterArea / 645.16M), 4, MidpointRounding.AwayFromZero).ToString() + " Sq in";
                            TensileStressArea_SI_Label.Text = Math.Round(MinorDiameterArea, 2, MidpointRounding.AwayFromZero).ToString() + " Sq mm";
                            TensileStressAreaTitleTab_Label.Text = "Minor Diameter Area";
                            TensileStressAreaValueTab_In_Label.Text = TensileStressArea_In_Label.Text;
                            TensileStressAreaValueTab_SI_Label.Text = TensileStressArea_SI_Label.Text;
                            Torque = Math.Round((ResidualStress * MinorDiameter * Coef_Friction), 0, MidpointRounding.AwayFromZero);
                            Torque_SI_Label.Text = Math.Round(Torque, 0, MidpointRounding.AwayFromZero).ToString() + " N-m";
                            TorqueTab_SI_Label.Text = Torque_SI_Label.Text;
                            Torque_In_Label.Text = Math.Round((Torque / 1.35581794833M), 0, MidpointRounding.AwayFromZero).ToString() + " ft-lbs";
                            TorqueValueTab_In_Label.Text = Torque_In_Label.Text;
                            Torque_SI_Sequence_Label.Text = Torque_SI_Label.Text;
                            Torque_In_Sequence_Label.Text = Torque_In_Label.Text;
                        }
                        else
                        {
                            MessageBox.Show("Pitch (No Threads Per Inch) value is wrong. Only non-zero numeral allowed", "Error!");
                        }
                    }
                    else
                    {
                        // Units are in inch here.
                        Bolt_Diameter = Bolt_Diameter.Substring(0, (Bolt_Diameter.Length - 1)).Trim();
                        NominalDiameter = Convert_Decimal(Bolt_Diameter);
                        if ((Pitch_TextBox.Text.Trim().Length > 0) && (Convert.ToDecimal(Pitch_TextBox.Text.Trim()) > 0M))
                        {
                            Pitch = Convert.ToDecimal(Pitch_TextBox.Text.Trim());
                            Pitch_Equivalent = 1M / Pitch;
                            MinorDiameter = NominalDiameter - ((10M / 16M) * Convert.ToDecimal(Math.Sqrt(3)) * Pitch_Equivalent);
                            MinorDiameterArea = ((22M / 7M) * Convert.ToDecimal(Math.Pow(Convert.ToDouble(MinorDiameter) / 2, 2)));
                            GlobalMinorDiameterArea = MinorDiameterArea;
                            TensileStressArea_Title_Label.Text = "Minor Diameter Area";
                            TensileStressArea_In_Label.Text = Math.Round(MinorDiameterArea, 4, MidpointRounding.AwayFromZero).ToString() + " Sq in";
                            TensileStressArea_SI_Label.Text = Math.Round((MinorDiameterArea * 645.16M), 2, MidpointRounding.AwayFromZero).ToString() + " Sq mm";
                            TensileStressAreaTitleTab_Label.Text = "Minor Diameter Area";
                            TensileStressAreaValueTab_In_Label.Text = TensileStressArea_In_Label.Text;
                            TensileStressAreaValueTab_SI_Label.Text = TensileStressArea_SI_Label.Text;
                            Torque = Math.Round((ResidualStress * (MinorDiameter * 25.4M) * Coef_Friction), 0, MidpointRounding.AwayFromZero);
                            Torque_SI_Label.Text = Math.Round((Torque), 0, MidpointRounding.AwayFromZero).ToString() + " N-m";
                            TorqueTab_SI_Label.Text = Torque_SI_Label.Text;
                            Torque_In_Label.Text = Math.Round((Torque / 1.35581794833M), 0, MidpointRounding.AwayFromZero).ToString() + " ft-lbs";
                            TorqueValueTab_In_Label.Text = Torque_In_Label.Text;
                            MinorDiameterArea = MinorDiameterArea * 645.16M;
                            Torque_SI_Sequence_Label.Text = Torque_SI_Label.Text;
                            Torque_In_Sequence_Label.Text = Torque_In_Label.Text;
                        }
                        else
                        {
                            MessageBox.Show("Threads Per Inch value is wrong. Only non-zero numeral allowed. Do not enter pitch.", "Error!");
                        }
                    }
                    OperateTool();
                }
            }
            else
            {
                if (TensileStressArea_RadioButton.Checked)
                {
                    TensileStressArea_Title_Label.Text = "Tensile Stress Area";
                    TensileStressArea_In_Label.Text = "";
                    TensileStressArea_SI_Label.Text = "";

                }
                else
                {
                    TensileStressArea_Title_Label.Text = "Minor Diameter Area";
                    TensileStressArea_In_Label.Text = "";
                    TensileStressArea_SI_Label.Text = "";
                }
            }
        }

        
 
        private void ImperialGraph_RadioButton_CheckedChanged(object sender, EventArgs e)
        {
            if (ImperialGraph_RadioButton.Checked)
            {
                DrawGraph();
            }
            else
            {
                DrawGraph();
            }
        }

        private void JointId_TextBox_Leave(object sender, EventArgs e)
        {
            Id_Label.Text = JointId_TextBox.Text;
        }

        
        protected void ManualCalculations()
        {
            /*
             * Manual input is very specific case where manually values are entered that are non-standard.
             * Following sequence is used to bring it to main stream
            */
            decimal Manual_LTF = 0M;
            decimal Manual_ResidualBoltStress = 0M;
            decimal Manual_Coefficient_Friction = 0M;
            decimal TopClamp = 0M;
            decimal BottomClamp = 0M;
            decimal GasketGap = 0M;
            decimal SpacerThickness = 0M;
            decimal ClampLength = 0M;
            string SelectedBoltDiameter = string.Empty;
            decimal BoltDiameter_for_StressArea = 0M;
            decimal Pitch = 0M;
            decimal StressArea = 0M;
            decimal Bolt_Length = 0M;
            decimal Torque = 0M;
            decimal pitch_equivalent = 0M;

            Manual_LTF = Calculate_LTF();

            // Info Panel
            LTF_Label.Text = Manual_LTF.ToString();

            if ((ResidualBoltStress_TextBox.Text.Trim() != "") && (ResidualBoltStress_TextBox.Text.Length > 0))
            {
                Manual_ResidualBoltStress = Convert.ToDecimal(ResidualBoltStress_TextBox.Text);
            }
            BoltMaterial_ComboBox.BackColor = Color.White;
            BoltThread_ComboBox.BackColor = Color.White;
            NoOfBoltsInfo_Label.Text = NumberOfBolts_TextBox.Text;
            BoltToolInfo_Label.Text = BoltTool_ComboBox.SelectedItem.ToString();
            if (Clamp1_TextBox.Text.Length > 0)
            {
                TopClamp = Convert.ToDecimal(Clamp1_TextBox.Text);
            }
            if (Clamp2_TextBox.Text.Length > 0)
            {
                BottomClamp = Convert.ToDecimal(Clamp2_TextBox.Text);
            }
            if (GasketGap_TextBox.Text.Length > 0)
            {
                GasketGap = Convert.ToDecimal(GasketGap_TextBox.Text);
            }
            if (CustomSpacer_TextBox.Text.Length > 0)
            {
                SpacerThickness = Convert.ToDecimal(CustomSpacer_TextBox.Text);
            }
            ClampLength = TopClamp + BottomClamp + GasketGap + SpacerThickness;

            if ((Coeff_TextBox.Text.Trim() != "") && (Coeff_TextBox.Text.Trim().Length > 0))
            {
                Manual_Coefficient_Friction = Convert.ToDecimal(Coeff_TextBox.Text.Trim());
            }

            if (UnitSystem_Button.Text == "SAE")
            {
                ClampLengthInfo_Label.Text = Math.Round(ClampLength, 4, MidpointRounding.AwayFromZero).ToString() + " in";
                ClampLengthInfo_mm_Label.Text = Math.Round((ClampLength * 25.4M), 2, MidpointRounding.AwayFromZero).ToString() + " mm";
            }
            else
            {
                ClampLengthInfo_Label.Text = (Math.Round((ClampLength / 25.4M), 4, MidpointRounding.AwayFromZero)).ToString() + " in";
                ClampLengthInfo_mm_Label.Text = Math.Round(ClampLength, 2, MidpointRounding.AwayFromZero).ToString() + " mm";
            }

            SelectedBoltDiameter = ((DataRowView)BoltThread_ComboBox.SelectedItem)["BoltThread"].ToString();

            if ((Pitch_TextBox.Text.Trim() != "") && (Pitch_TextBox.Text.Length > 0))
            {
                try
                {
                    Pitch = Convert.ToDecimal(Pitch_TextBox.Text.Trim());
                }
                catch (Exception Pitch_Exception)
                {
                    MessageBox.Show(Pitch_Exception.Message + " during manual setup.");
                }

                if (SelectedBoltDiameter.Substring(0, 1) == "M")
                {
                    BoltDiameter_for_StressArea = Convert.ToDecimal(SelectedBoltDiameter.Substring(1));
                    StressArea = 0.7854M * (BoltDiameter_for_StressArea - 0.938194M * Pitch) * (BoltDiameter_for_StressArea - 0.938194M * Pitch);
                    TensileStressArea_SI_Label.Text = Math.Round(StressArea, 2, MidpointRounding.AwayFromZero).ToString() + " Sq mm";
                    TensileStressArea_In_Label.Text = Math.Round(Convert.ToDecimal((StressArea / 645.16M)), 4, MidpointRounding.AwayFromZero).ToString() + " Sq in";
                    TensileStressAreaValueTab_In_Label.Text = TensileStressArea_In_Label.Text;
                    TensileStressAreaValueTab_SI_Label.Text = TensileStressArea_SI_Label.Text;
                    if (UnitSystem_Button.Text == "SAE")
                    {
                        Bolt_Length = Math.Round((ClampLength + 3 * (BoltDiameter_for_StressArea / 25.4M) + 3 * (Pitch / 25.4M)), 4, MidpointRounding.AwayFromZero);
                        BoltLength_In_Label.Text = Bolt_Length.ToString() + " in";
                        BoltLength_SI_Label.Text = Math.Round((Bolt_Length * 25.4M), 2, MidpointRounding.AwayFromZero).ToString() + " mm";
                        BoltLengthValueTab_In_Label.Text = BoltLength_In_Label.Text;
                        BoltLengthValueTab_SI_Label.Text = BoltLength_SI_Label.Text;
                        Torque = (Manual_ResidualBoltStress / 145.037738007M) * Manual_Coefficient_Friction * BoltDiameter_for_StressArea;
                        Torque_In_Label.Text = Math.Round((Torque / 1.35581794833M), 0, MidpointRounding.AwayFromZero).ToString() + " lb-ft";
                        Torque_SI_Label.Text = Math.Round(Torque, 0, MidpointRounding.AwayFromZero).ToString() + " N-m";
                        TorqueTab_SI_Label.Text = Torque_SI_Label.Text;
                        TorqueValueTab_In_Label.Text = Torque_In_Label.Text;
                        Torque_SI_Sequence_Label.Text = Torque_SI_Label.Text;
                        Torque_In_Sequence_Label.Text = Torque_In_Label.Text;
                    }
                    else
                    {
                        Bolt_Length = ClampLength + 3 * BoltDiameter_for_StressArea + 3 * Pitch;
                        BoltLength_SI_Label.Text = Math.Round(Bolt_Length, 2, MidpointRounding.AwayFromZero).ToString() + " mm";
                        BoltLength_In_Label.Text = Math.Round((Bolt_Length / 25.4M), 4, MidpointRounding.AwayFromZero).ToString() + " in";
                        BoltLengthValueTab_In_Label.Text = BoltLength_In_Label.Text;
                        BoltLengthValueTab_SI_Label.Text = BoltLength_SI_Label.Text;
                        Torque = (Manual_ResidualBoltStress * Manual_Coefficient_Friction * BoltDiameter_for_StressArea);
                        Torque_In_Label.Text = Math.Round((Torque * 0.737562149277M), 0, MidpointRounding.AwayFromZero).ToString() + " lb-ft";
                        Torque_SI_Label.Text = Math.Round(Torque, 0, MidpointRounding.AwayFromZero).ToString() + " N-m";
                        TorqueTab_SI_Label.Text = Torque_SI_Label.Text;
                        TorqueValueTab_In_Label.Text = Torque_In_Label.Text;
                        Torque_SI_Sequence_Label.Text = Torque_SI_Label.Text;
                        Torque_In_Sequence_Label.Text = Torque_In_Label.Text;
                    }
                }
                else
                {
                    SelectedBoltDiameter = SelectedBoltDiameter.Substring(0, (SelectedBoltDiameter.Length - 1)).Trim();
                    BoltDiameter_for_StressArea = Convert_Decimal(SelectedBoltDiameter);
                    BoltDiameter_for_StressArea = Math.Round((BoltDiameter_for_StressArea * 25.4M), 4, MidpointRounding.AwayFromZero);
                    StressArea = 0.7854M * (BoltDiameter_for_StressArea - (0.9743M / Pitch)) * (BoltDiameter_for_StressArea - (0.9743M / Pitch));
                    TensileStressArea_SI_Label.Text = Math.Round(StressArea, 2, MidpointRounding.AwayFromZero).ToString() + " Sq mm";
                    TensileStressArea_In_Label.Text = Math.Round(Convert.ToDecimal((StressArea / 645.16M)), 4, MidpointRounding.AwayFromZero).ToString() + " Sq in";
                    TensileStressAreaValueTab_In_Label.Text = TensileStressArea_In_Label.Text;
                    TensileStressAreaValueTab_SI_Label.Text = TensileStressArea_SI_Label.Text;
                    pitch_equivalent = 1 / Pitch;

                    if (UnitSystem_Button.Text == "SAE")
                    {

                        Bolt_Length = ClampLength + 3 * (BoltDiameter_for_StressArea / 25.4M) + 3 * pitch_equivalent;
                        BoltLength_In_Label.Text = Math.Round(Bolt_Length, 4, MidpointRounding.AwayFromZero).ToString() + " in";
                        BoltLength_SI_Label.Text = Math.Round((Bolt_Length * 25.4M), 2, MidpointRounding.AwayFromZero).ToString() + " mm";
                        BoltLengthValueTab_In_Label.Text = BoltLength_In_Label.Text;
                        BoltLengthValueTab_SI_Label.Text = BoltLength_SI_Label.Text;
                        Torque = (Manual_ResidualBoltStress / 145.037738007M) * Manual_Coefficient_Friction * BoltDiameter_for_StressArea;
                        Torque_In_Label.Text = Math.Round((Torque / 1.35581794833M), 0, MidpointRounding.AwayFromZero).ToString() + " lb-ft";
                        Torque_SI_Label.Text = Math.Round(Torque, 0, MidpointRounding.AwayFromZero).ToString() + " N-m";
                        TorqueTab_SI_Label.Text = Torque_SI_Label.Text;
                        TorqueValueTab_In_Label.Text = Torque_In_Label.Text;
                        Torque_SI_Sequence_Label.Text = Torque_SI_Label.Text;
                        Torque_In_Sequence_Label.Text = Torque_In_Label.Text;
                    }
                    else
                    {
                        Bolt_Length = ClampLength + 3 * BoltDiameter_for_StressArea + 3 * pitch_equivalent * 25.4M;
                        BoltLength_SI_Label.Text = Math.Round(Bolt_Length, 2, MidpointRounding.AwayFromZero).ToString() + " mm";
                        BoltLength_In_Label.Text = Math.Round((Bolt_Length / 25.4M), 4, MidpointRounding.AwayFromZero).ToString() + " in";
                        BoltLengthValueTab_In_Label.Text = BoltLength_In_Label.Text;
                        BoltLengthValueTab_SI_Label.Text = BoltLength_SI_Label.Text;
                        Torque = (Manual_ResidualBoltStress * Manual_Coefficient_Friction * BoltDiameter_for_StressArea);
                        Torque_In_Label.Text = Math.Round((Torque * 0.737562149277M), 0, MidpointRounding.AwayFromZero).ToString() + " lb-ft";
                        Torque_SI_Label.Text = Math.Round(Torque, 0, MidpointRounding.AwayFromZero).ToString() + " N-m";
                        TorqueTab_SI_Label.Text = Torque_SI_Label.Text;
                        TorqueValueTab_In_Label.Text = Torque_In_Label.Text;
                        Torque_SI_Sequence_Label.Text = Torque_SI_Label.Text;
                        Torque_In_Sequence_Label.Text = Torque_In_Label.Text;
                    }
                }
            }
            // End of Info Panel

            // Bolt Stress Tab and Torque Tab is covered during Info Panel

            // Graph Tab
            DrawGraph();

            // Bolt Tab
            NominalBoltDiameterValueTab_Label.Text = ((DataRowView)BoltThread_ComboBox.SelectedItem)["BoltThread"].ToString();
            NumberOfBoltsValueTab_Label.Text = NumberOfBolts_TextBox.Text;

            // Sequence Panel
            int Bolts = Convert.ToInt32(NumberOfBolts_TextBox.Text);
            switch (Bolts)
            {
                case 4:
                    Sequence_PictureBox.Image = WizBolt.Properties.Resources.FlangeTorque_04;
                    break;
                case 8:
                    Sequence_PictureBox.Image = WizBolt.Properties.Resources.FlangeTorque_08;
                    break;
                case 12:
                    Sequence_PictureBox.Image = WizBolt.Properties.Resources.FlangeTorque_12;
                    break;
                case 16:
                    Sequence_PictureBox.Image = WizBolt.Properties.Resources.FlangeTorque_16;
                    break;
                case 20:
                    Sequence_PictureBox.Image = WizBolt.Properties.Resources.FlangeTorque_20;
                    break;
                case 24:
                    Sequence_PictureBox.Image = WizBolt.Properties.Resources.FlangeTorque_24;
                    break;
                case 28:
                    Sequence_PictureBox.Image = WizBolt.Properties.Resources.FlangeTorque_28;
                    break;
                case 32:
                    Sequence_PictureBox.Image = WizBolt.Properties.Resources.FlangeTorque_32;
                    break;
                case 36:
                    Sequence_PictureBox.Image = WizBolt.Properties.Resources.FlangeTorque_36;
                    break;
                case 40:
                    Sequence_PictureBox.Image = WizBolt.Properties.Resources.FlangeTorque_40;
                    break;
                case 44:
                    Sequence_PictureBox.Image = WizBolt.Properties.Resources.FlangeTorque_44;
                    break;
                case 48:
                    Sequence_PictureBox.Image = WizBolt.Properties.Resources.FlangeTorque_48;
                    break;
                case 52:
                    Sequence_PictureBox.Image = WizBolt.Properties.Resources.FlangeTorque_52;
                    break;
                case 56:
                    Sequence_PictureBox.Image = WizBolt.Properties.Resources.FlangeTorque_56;
                    break;
                case 60:
                    Sequence_PictureBox.Image = WizBolt.Properties.Resources.FlangeTorque_60;
                    break;
                default:
                    Sequence_PictureBox.Image = null;
                    break;
            }

            IsManual = false;
            BoltTool_ComboBox.Enabled = true;
        }

        private void ResidualBoltStress_TextBox_Leave(object sender, EventArgs e)
        {

            if (!IsInitial)
            {
                decimal ResidualBoltStress = 0M;
                decimal YieldPC = 0M;
                decimal BoltYield = 0M;
                decimal LTFValue = 0M;
                decimal StressArea = 0M;
                decimal ToolPressureArea = 0M;
                decimal CrossLoading = 0M;
                decimal Coefficient_Friction = 0M;
                decimal Torque = 0M;
                decimal BoltDiameter_for_StressArea = 0M;
                decimal Pitch = 0M;
                decimal pitch_equivalent = 0M;
                string Bolt_Diameter = string.Empty;
                if (!IsSetting)
                {
                    // Switch off the warnings. These are displayed if it is required at respective places.
                    TitleWarning_PictureBox.Visible = false;

                    PressureAWarning_PictureBox.Visible = false;
                    PressureBWarning_PictureBox.Visible = false;
                    ResidualStreeWarning_PictureBox.Visible = false;

                    PressureA_Info_Warning_PictureBox.Visible = false;
                    PressureB_Info_Warning_PictureBox.Visible = false;
                    Deten_Info_Warning_PictureBox.Visible = false;
                    if (!IsLoading)
                    {
                        if ((BoltYield_SI_Label.Text.Trim() != "") && (BoltYield_SI_Label.Text.Trim().Length > 4))
                        {
                            BoltYield = Convert.ToDecimal(BoltYield_SI_Label.Text.Trim().Substring(0, BoltYield_SI_Label.Text.Trim().Length - 4));
                        }
                        if ((LTF_Label.Text.Trim() != "") && (LTF_Label.Text.Trim().Length > 0))
                        {
                            LTFValue = Convert.ToDecimal(LTF_Label.Text.Trim());
                        }
                        else
                        {
                            LTFValue = 1.11M;
                        }
                        if ((TensileStressArea_SI_Label.Text.Trim() != "") && (TensileStressArea_SI_Label.Text.Trim().Length > 6))
                        {
                            StressArea = Convert.ToDecimal(TensileStressArea_SI_Label.Text.Trim().Substring(0, TensileStressArea_SI_Label.Text.Trim().Length - 6));
                        }
                        if ((ToolPressureArea_mm_Label.Text.Trim() != "") && (ToolPressureArea_mm_Label.Text.Trim().Length > 6))
                        {
                            ToolPressureArea = Convert.ToDecimal(ToolPressureArea_mm_Label.Text.Trim().Substring(0, ToolPressureArea_mm_Label.Text.Trim().Length - 6));
                        }
                        if ((CrossLoading_TextBox.Text.Trim() != "") && (CrossLoading_TextBox.Text.Trim().Length > 0))
                        {
                            CrossLoading = Convert.ToDecimal(CrossLoading_TextBox.Text.Trim());
                        }
                        if ((CoefficientValue_Label.Text.Trim() != "") && (CoefficientValue_Label.Text.Trim().Length > 0))
                        {
                            Coefficient_Friction = Convert.ToDecimal(CoefficientValue_Label.Text.Trim());
                        }
                        if ((Pitch_TextBox.Text.Trim() != "") && (Pitch_TextBox.Text.Trim().Length > 0))
                        {
                            Pitch = Convert.ToDecimal(Pitch_TextBox.Text.Trim());

                        }
                        if (Pitch > 0)
                        {
                            pitch_equivalent = 1.0M / Pitch;
                        }
                        try
                        {
                            if (UnitSystem_Button.Text == "SAE")
                            {
                                if ((ResidualBoltStress_TextBox.Text != "") && (ResidualBoltStress_TextBox.Text.Trim().Length > 0))
                                {
                                    if (Convert.ToDecimal(ResidualBoltStress_TextBox.Text.Trim()) > 0)
                                    {
                                        {
                                            ResidualBoltStress = Convert.ToDecimal(ResidualBoltStress_TextBox.Text.Trim());
                                            ResidualBoltStress = Math.Round((ResidualBoltStress / 145.037738007M), 2, MidpointRounding.AwayFromZero);
                                        }
                                    }
                                }
                            }
                            else
                            {
                                if ((ResidualBoltStress_TextBox.Text != "") && (ResidualBoltStress_TextBox.Text.Trim().Length > 0))
                                {
                                    if (Convert.ToDecimal(ResidualBoltStress_TextBox.Text.Trim()) > 0)
                                    {
                                        ResidualBoltStress = Convert.ToDecimal(ResidualBoltStress_TextBox.Text.Trim());
                                    }
                                }
                            }

                            YieldPC = ((ResidualBoltStress * LTFValue * 100) / BoltYield);
                            if (YieldPC >= 95)
                            {
                                MessageBox.Show("This value of Residual Bolt Stress exceeds 95% of Yield and hence not acceptable. Plan for lesser value.");
                                ResidualBoltStress_TextBox.Text = "";
                              //  DefaultStress_Button.Visible = true;
                            }
                            else
                            {
                                // Torque 
                                Bolt_Diameter = ((DataRowView)BoltThread_ComboBox.SelectedItem)["BoltThread"].ToString();
                                if (Bolt_Diameter.Substring(0, 1) == "M")
                                {
                                    BoltDiameter_for_StressArea = Convert.ToDecimal(Bolt_Diameter.Substring(1));
                                    //Torque = (ResidualBoltStress * Coefficient_Friction * BoltDiameter_for_StressArea);
                                    Torque = ResidualBoltStress * ((Pitch / (2M * 3.14159M)) + (Coefficient_Friction * BoltDiameter_for_StressArea) / 1.7320508M + (Coefficient_Friction * BoltDiameter_for_StressArea * 1.5M / 2M));
                                }
                                else
                                {
                                    Bolt_Diameter = Bolt_Diameter.Substring(0, (Bolt_Diameter.Length - 1)).Trim();
                                    BoltDiameter_for_StressArea = Convert_Decimal(Bolt_Diameter);
                                    // Torque = (ResidualBoltStress * Coefficient_Friction * BoltDiameter_for_StressArea * 25.4M);
                                    Torque = ResidualBoltStress * ((pitch_equivalent * 25.4M / (2M * 3.14159M)) + (Coefficient_Friction * BoltDiameter_for_StressArea * 25.4M) / 1.7320508M + (Coefficient_Friction * BoltDiameter_for_StressArea * 25.4M * 1.5M / 2M));
                                }
                                Torque_In_Label.Text = Math.Round((Torque * 0.737562149277M), 0, MidpointRounding.AwayFromZero).ToString() + " lb-ft";
                                Torque_SI_Label.Text = Math.Round(Torque, 0, MidpointRounding.AwayFromZero).ToString() + " N-m";
                                TorqueTab_SI_Label.Text = Math.Round(Torque, 0, MidpointRounding.AwayFromZero).ToString() + " N-m";
                                TorqueValueTab_In_Label.Text = Math.Round((Torque * 0.737562149277M), 0, MidpointRounding.AwayFromZero).ToString() + " lb-ft";
                                Torque_SI_Sequence_Label.Text = Torque_SI_Label.Text;
                                Torque_In_Sequence_Label.Text = Torque_In_Label.Text;

                                OperateTool();
                            }
                        }
                        catch (Exception ResidualStress_Exceprion)
                        {
                            MessageBox.Show(ResidualStress_Exceprion.Message + " while Residual stress is changed.");
                        }
                        if (IsManual)
                        {
                            if ((ResidualBoltStress_TextBox.Text.Trim() != "") && (ResidualBoltStress_TextBox.Text.Length > 0))
                            {
                                try
                                {
                                    decimal Stress = Convert.ToDecimal(ResidualBoltStress_TextBox.Text.Trim());
                                    ResidualBoltStress_TextBox.BackColor = Color.White;
                                    if ((Clamp2_TextBox.Text.Trim() != "") && (Clamp2_TextBox.Text.Length > 0)
                                        && (Clamp1_TextBox.Text.Trim() != "") && (Clamp1_TextBox.Text.Length > 0)
                                        && (NumberOfBolts_TextBox.Text.Trim() != "") && (NumberOfBolts_TextBox.Text.Length > 0))
                                    {
                                        ManualCalculations();
                                    }
                                }
                                catch (Exception ResidualBoltStress_Exception)
                                {
                                    MessageBox.Show(ResidualBoltStress_Exception.Message + " only numeric value accepted. Error occurred during manual input of Residual Bolt Stress.");
                                }

                            }
                        }
                        // Graph
                        DrawGraph();
                    }
                }
            }
        }

        private void Clamp1_TextBox_Leave(object sender, EventArgs e)
        {
            if (IsManual)
            {

                if ((Clamp1_TextBox.Text.Trim() != "") && (Clamp1_TextBox.Text.Length > 0))
                {
                    try
                    {
                        decimal ClampLength = Convert.ToDecimal(Clamp1_TextBox.Text.Trim());
                        Clamp1_TextBox.BackColor = Color.White;
                        if ((Clamp2_TextBox.Text.Trim() != "") && (Clamp2_TextBox.Text.Length > 0)
                            && (ResidualBoltStress_TextBox.Text.Trim() != "") && (ResidualBoltStress_TextBox.Text.Length > 0)
                            && (NumberOfBolts_TextBox.Text.Trim() != "") && (NumberOfBolts_TextBox.Text.Length > 0))
                        {
                            ManualCalculations();
                        }
                    }
                    catch (Exception Clamp1_Exception)
                    {
                        MessageBox.Show(Clamp1_Exception.Message + " only numeric value accepted. Error occurred during manual input of upper flange.");
                    }

                }
            }

        }

        private void Clamp2_TextBox_Leave(object sender, EventArgs e)
        {
            if (IsManual)
            {

                if ((Clamp2_TextBox.Text.Trim() != "") && (Clamp2_TextBox.Text.Length > 0))
                {
                    try
                    {
                        decimal ClampLength = Convert.ToDecimal(Clamp2_TextBox.Text.Trim());
                        Clamp2_TextBox.BackColor = Color.White;
                        if ((Clamp1_TextBox.Text.Trim() != "") && (Clamp1_TextBox.Text.Length > 0)
                            && (ResidualBoltStress_TextBox.Text.Trim() != "") && (ResidualBoltStress_TextBox.Text.Length > 0)
                            && (NumberOfBolts_TextBox.Text.Trim() != "") && (NumberOfBolts_TextBox.Text.Length > 0))
                        {
                            ManualCalculations();
                        }
                    }
                    catch (Exception Clamp2_Exception)
                    {
                        MessageBox.Show(Clamp2_Exception.Message + " only numeric value accepted. Error occurred during manual input of lower flange.");
                    }

                }
            }
        }

        private void NumberOfBolts_TextBox_Leave(object sender, EventArgs e)
        {
            if (IsManual)
            {

                if ((NumberOfBolts_TextBox.Text.Trim() != "") && (NumberOfBolts_TextBox.Text.Length > 0))
                {
                    try
                    {
                        decimal Bolts = Convert.ToDecimal(NumberOfBolts_TextBox.Text.Trim());
                        NumberOfBolts_TextBox.BackColor = Color.White;
                        if ((Clamp2_TextBox.Text.Trim() != "") && (Clamp2_TextBox.Text.Length > 0)
                            && (ResidualBoltStress_TextBox.Text.Trim() != "") && (ResidualBoltStress_TextBox.Text.Length > 0)
                            && (Clamp1_TextBox.Text.Trim() != "") && (Clamp1_TextBox.Text.Length > 0))
                        {
                            ManualCalculations();
                        }
                    }
                    catch (Exception BoltNo_Exception)
                    {
                        MessageBox.Show(BoltNo_Exception.Message + " only numeric value accepted. Error occurred during manual input of bolts.");
                    }

                }
            }
        }

        protected void DrawGraph()
        {
            try
            {
                // Create Graph
                ChartArea StressChartArea = Stress_Chart.ChartAreas.FindByName("StressChart");
                Series BoltYield_Series = Stress_Chart.Series.FindByName("Bolt Yield Stress");
                Series Detension_Series = Stress_Chart.Series.FindByName("De-tensioning Stress");
                Series APressure_Series = Stress_Chart.Series.FindByName("A Pressure Stress");
                Series BPressure_Series = Stress_Chart.Series.FindByName("B Pressure Stress");
                StressChartArea.AxisX.MajorGrid.LineDashStyle = System.Windows.Forms.DataVisualization.Charting.ChartDashStyle.DashDotDot;
                StressChartArea.AxisY.MajorGrid.LineDashStyle = System.Windows.Forms.DataVisualization.Charting.ChartDashStyle.DashDotDot;
                // 
                // Clear the previous
                BoltYield_Series.Points.Clear();
                Detension_Series.Points.Clear();
                APressure_Series.Points.Clear();
                BPressure_Series.Points.Clear();
                Double BoltYield_Y = 0;
                Double Detension_Y = 0;
                Double APressure_Y = 0;
                Double ResidualPressure_Y = 0;
                Double BPressure_Y = 0;

                if ((BoltYield_In_Label.Text != "") && (BoltYield_In_Label.Text.Length > 4))
                {
                    if (ImperialGraph_RadioButton.Checked)
                    {
                        BoltYield_Y = Convert.ToDouble(BoltYield_In_Label.Text.Substring(0, BoltYield_In_Label.Text.Length - 4)) * 1000;
                    }
                    else
                    {
                        BoltYield_Y = Convert.ToDouble(BoltYield_In_Label.Text.Substring(0, BoltYield_In_Label.Text.Length - 4)) * 1000 / 145.037738007;
                    }
                }
                
               
                if ((T3RBoltStress_In_Label.Text != "") && (T3RBoltStress_In_Label.Text.Length > 0))
                {
                    if (Convert.ToDouble(T3RBoltStress_In_Label.Text) > 0)
                    {
                        if (ImperialGraph_RadioButton.Checked)
                        {
                            ResidualPressure_Y = Convert.ToDouble(T3RBoltStress_In_Label.Text);
                        }
                        else
                        {
                            ResidualPressure_Y = Convert.ToDouble(T3RBoltStress_In_Label.Text) / 145.037738007;
                        }
                    }
                }

                // Bolt Yield Graph
                BoltYield_Series.XValueType = ChartValueType.Int32;
                BoltYield_Series.Points.Add(new DataPoint(0, BoltYield_Y));
                BoltYield_Series.Points.Add(new DataPoint(3, BoltYield_Y));
                // Following also works
                //BoltYield_Series.Points.AddXY(0, BoltYield_Y);
                //BoltYield_Series.Points.AddXY(3, BoltYield_Y);
                BoltYield_Series.ChartType = SeriesChartType.Line;

                if ((DetenBoltStress_In_Label.Text != "") && (DetenBoltStress_In_Label.Text.Length > 0))
                {
                    if (Convert.ToDouble(DetenBoltStress_In_Label.Text) > 0)
                    {
                        if (ImperialGraph_RadioButton.Checked)
                        {
                            Detension_Y = Convert.ToDouble(DetenBoltStress_In_Label.Text);
                        }
                        else
                        {
                            Detension_Y = Convert.ToDouble(DetenBoltStress_In_Label.Text) / 145.037738007;
                        }
                        Detension_Series.XValueType = ChartValueType.Int32;

                        Detension_Series.Points.Add(new DataPoint(0, Detension_Y));
                        Detension_Series.Points.Add(new DataPoint(3, Detension_Y));
                        // Following also works
                        //Detension_Series.Points.AddXY(0, Detension_Y);
                        //Detension_Series.Points.AddXY(3, Detension_Y);
                        Detension_Series.ChartType = SeriesChartType.Line;
                    }
                }

                // A Pressure
                if ((T1ABoltStress_In_Label.Text != "") && (T1ABoltStress_In_Label.Text.Length > 0))
                {
                    if (Convert.ToDouble(T1ABoltStress_In_Label.Text) > 0)
                    {
                        if (ImperialGraph_RadioButton.Checked)
                        {
                            APressure_Y = Convert.ToDouble(T1ABoltStress_In_Label.Text);
                        }
                        else
                        {
                            APressure_Y = Convert.ToDouble(T1ABoltStress_In_Label.Text) / 145.037738007;
                        }
                        APressure_Series.XValueType = ChartValueType.Int32;
                        APressure_Series.Points.Add(new DataPoint(0, 0));
                        APressure_Series.Points.Add(new DataPoint(1, APressure_Y));
                        APressure_Series.Points.Add(new DataPoint(2, ResidualPressure_Y));
                        APressure_Series.Points.Add(new DataPoint(3, ResidualPressure_Y));
                        APressure_Series.ChartType = SeriesChartType.Line;
                    }
                }

                // B Pressure
                if ((T1BBoltStress_In_Label.Text != "") && (T1BBoltStress_In_Label.Text.Length > 0))
                {
                    if (Convert.ToDouble(T1BBoltStress_In_Label.Text) > 0)
                    {
                        if (ImperialGraph_RadioButton.Checked)
                        {
                            BPressure_Y = Convert.ToDouble(T1BBoltStress_In_Label.Text);
                        }
                        else
                        {
                            BPressure_Y = Convert.ToDouble(T1BBoltStress_In_Label.Text) / 145.037738007;
                        }
                       
                    }
                }


                // Residual Pressure
                if ((ResidualPressure_Y > 0) && (BPressure_Y > 0))
                {
                    BPressure_Series.XValueType = ChartValueType.Int32;
                    BPressure_Series.Points.Add(new DataPoint(0, 0));
                    BPressure_Series.Points.Add(new DataPoint(1, BPressure_Y));
                    BPressure_Series.Points.Add(new DataPoint(2, ResidualPressure_Y));
                    BPressure_Series.Points.Add(new DataPoint(3, ResidualPressure_Y));
                    BPressure_Series.ChartType = SeriesChartType.Line;
                }
                
            }
            catch (Exception Graph_Exception)
            {
                MessageBox.Show(Graph_Exception.Message + " when drawing graph.");
            }
        }
        public void SpecialTools()
        {
            SpecialToolForm.ShowDialog();
        }

        private void TensionerTool_ComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            
                decimal Hydraulic_Area = 0M;
                decimal ToolPressureB = 0M;
                decimal ToolPressureA = 0M;
                decimal StressArea = 0M;
                decimal ResidualBoltStress = 0M;
                decimal PressureRelease = 0M;      // Pressure Release ratio is another name for Load Tranfer Ratio i.e. LTF
                decimal CrossLoad = 0M;


                    DataTable Tool_Table = (DataTable)TensionerTool_ComboBox.DataSource;
                    if (Tool_Table != null)
                    {
                        if (TensionerTool_ComboBox.SelectedIndex >= 0)
                        {
                            if ((Tool_Table.Rows[TensionerTool_ComboBox.SelectedIndex]["HydraulicArea"].ToString().Trim() != "") && (Tool_Table.Rows[TensionerTool_ComboBox.SelectedIndex]["HydraulicArea"].ToString().Trim().Length > 0))
                            {
                                Hydraulic_Area = Convert.ToDecimal(Tool_Table.Rows[TensionerTool_ComboBox.SelectedIndex]["HydraulicArea"]);
                            }
                            // Info Panel

                            ToolId_Info_Label.Text = ((DataRowView)TensionerTool_ComboBox.SelectedItem)["ModelNumber"].ToString();

                            ToolPressureArea_Info_Label.Text = Math.Round((Hydraulic_Area / 645.16M), 4, MidpointRounding.AwayFromZero).ToString() + " Sq in";
                            ToolPressureArea_mm_Label.Text = Math.Round(Hydraulic_Area, 0, MidpointRounding.AwayFromZero).ToString() + " Sq mm";

                            // Sequence Tab

                            if ((ResidualBoltStress_TextBox.Text.Trim() != "") && (ResidualBoltStress_TextBox.Text.Trim().Length > 0))
                            {
                                ResidualBoltStress = Convert.ToDecimal(ResidualBoltStress_TextBox.Text.Trim());
                            }
                            if (UnitSystem_Button.Text == "SAE")
                            {
                                ResidualBoltStress = ResidualBoltStress / 145.037738007M;
                            }
                            if ((LTF_Label.Text.Trim() != "") && (LTF_Label.Text.Trim().Length > 0))
                            {
                                PressureRelease = Convert.ToDecimal(LTF_Label.Text.Trim());
                            }

                            if (Hydraulic_Area > 0)
                            {
                                ToolPressureB = ResidualBoltStress * StressArea * PressureRelease / Hydraulic_Area;
                            }
                            if ((CrossLoading_TextBox.Text.Trim() != "") && (CrossLoading_TextBox.Text.Trim().Length > 0))
                            {
                                CrossLoad = Convert.ToDecimal(CrossLoading_TextBox.Text.Trim());
                            }
                            ToolPressureA = ToolPressureB * (1 + CrossLoad / 100M);

                            // Sequence Tab with embedded Info Panel
                            OperateTool();
                        }
                    }

        }

        protected void OperateTool()
        {
            // Local Variables
            decimal TensioningPressureB = 0M;
            decimal TensioningPressureA = 0M;
            decimal ResidualBoltStress = 0M;
            decimal TensileStressArea_in = 0M;
            decimal TensileStressArea_mm = 0M;
            decimal BoltYield = 0M;
            decimal LTFValue = 1.11M;
            decimal ToolHydraulicArea = 0M;
            decimal DetensionProportion = 0M;
            decimal CrossLoad = 0M;
            string ToolArea = string.Empty;

            // Switch off the warnings. These are displayed if it is required at respective places.
            TitleWarning_PictureBox.Visible = false;

            PressureAWarning_PictureBox.Visible = false;
            PressureBWarning_PictureBox.Visible = false;
            ResidualStreeWarning_PictureBox.Visible = false;

            PressureA_Info_Warning_PictureBox.Visible = false;
            PressureB_Info_Warning_PictureBox.Visible = false;
            Deten_Info_Warning_PictureBox.Visible = false;

            // BugFix: Even after changing the proportion (100%, 50%, or Torque) the Info Panel did not reflect it.
            if (BoltTool_ComboBox.SelectedItem != null)
            {
                BoltToolInfo_Label.Text = BoltTool_ComboBox.SelectedItem.ToString();
            }
            else
            {
                BoltTool_ComboBox.SelectedValue = "100%";
            }

            if ((ResidualBoltStress_TextBox.Text.Trim() != "") && (ResidualBoltStress_TextBox.Text.Trim().Length > 0))
            {
                ResidualBoltStress = Convert.ToDecimal(ResidualBoltStress_TextBox.Text.Trim());
            }
            if ((TensileStressArea_In_Label.Text.Trim() != "") && (TensileStressArea_In_Label.Text.Trim().Length > 5))
            {
                string StressAreaValue_in = TensileStressArea_In_Label.Text.Substring(0, TensileStressArea_In_Label.Text.Length - 5);
                TensileStressArea_in = Convert.ToDecimal(StressAreaValue_in);
            }
            if ((TensileStressArea_SI_Label.Text.Trim() != "") && (TensileStressArea_SI_Label.Text.Trim().Length > 5))
            {
                string StressAreaValue_mm = TensileStressArea_SI_Label.Text.Substring(0, TensileStressArea_SI_Label.Text.Length - 5);
                TensileStressArea_mm = Convert.ToDecimal(StressAreaValue_mm);
            }
            if ((BoltYield_SI_Label.Text.Trim() != "") && (BoltYield_SI_Label.Text.Trim().Length > 3))
            {
                string YieldValue = BoltYield_SI_Label.Text.Substring(0, BoltYield_SI_Label.Text.Length - 3);
                BoltYield = Convert.ToDecimal(YieldValue);
            }
            if ((LTF_Label.Text.Trim() != "") && (LTF_Label.Text.Trim().Length > 0))
            {
                LTFValue = Convert.ToDecimal(LTF_Label.Text.Trim());
            }
            if ((ToolPressureArea_mm_Label.Text.Trim() != "") && (ToolPressureArea_mm_Label.Text.Trim().Length > 6))
            {
                ToolArea = ToolPressureArea_mm_Label.Text.Substring(0, ToolPressureArea_mm_Label.Text.Length - 6);
                ToolHydraulicArea = Convert.ToDecimal(ToolArea);
            }
            
            if ((Detensioning_TextBox.Text.Trim() != "") && (Detensioning_TextBox.Text.Trim().Length > 0))
            {
                DetensionProportion = Convert.ToDecimal(Detensioning_TextBox.Text.Trim());
            }
            if ((CrossLoading_TextBox.Text.Trim() != "") && (CrossLoading_TextBox.Text.Trim().Length > 0))
            {
                CrossLoad = Convert.ToDecimal(CrossLoading_TextBox.Text.Trim());
            }
            int Bolt_Tool_Proportion = BoltTool_ComboBox.SelectedIndex;
            switch (Bolt_Tool_Proportion)
            {
                case 0:
                    Sequence_PictureBox.Image = WizBolt.Properties.Resources.Sequence_Full;
                    if (UnitSystem_Button.Text == "SAE")
                    {
                        if (ToolHydraulicArea > 0)
                        {
                            // Bolt Stress Tab Calculations for Bolt
                            T1BBoltStress_In_Label.Text = Math.Round((ResidualBoltStress * LTFValue), 0, MidpointRounding.AwayFromZero).ToString();
                            T1BBoltStress_SI_Label.Text = Math.Round((ResidualBoltStress * LTFValue / 145.037738007M), 2, MidpointRounding.AwayFromZero).ToString();
                            T1BBoltLoad_In_Label.Text = Math.Round((ResidualBoltStress * LTFValue * TensileStressArea_in / 2240M), 2, MidpointRounding.AwayFromZero).ToString();
                            T1BBoltLoad_SI_Label.Text = Math.Round((ResidualBoltStress * TensileStressArea_mm * LTFValue / 145037.738007M), 2, MidpointRounding.AwayFromZero).ToString();
                            if (BoltYield > 0)
                            {
                                T1BBoltYield_Per_Label.Text = Math.Round(((ResidualBoltStress * LTFValue / 145.037738007M) * 100M / BoltYield), 2, MidpointRounding.AwayFromZero).ToString();
                            }
                            else
                            {
                                T1BBoltYield_Per_Label.Text = "";
                            }

                            T1ABoltStress_In_Label.Text = "";
                            T1ABoltStress_SI_Label.Text = "";
                            T1ABoltLoad_In_Label.Text = "";
                            T1ABoltLoad_SI_Label.Text = "";
                            T1ABoltYield_SI_Label.Text = "";

                            T2RBoltStress_SI_Label.Text = Math.Round((ResidualBoltStress / 145.037738007M), 2, MidpointRounding.AwayFromZero).ToString();
                            T2RBoltLoad_In_Label.Text = Math.Round((ResidualBoltStress * TensileStressArea_in / 2240M), 2, MidpointRounding.AwayFromZero).ToString();
                            T2RBoltLoad_SI_Label.Text = Math.Round((ResidualBoltStress * TensileStressArea_mm / 145037.738007M), 2, MidpointRounding.AwayFromZero).ToString();
                            if (BoltYield > 0)
                            {
                                T2RBoltYield_Per_Label.Text = Math.Round(((ResidualBoltStress / 145.037738007M) * 100M / BoltYield), 2, MidpointRounding.AwayFromZero).ToString();
                            }
                            else
                            {
                                T2RBoltYield_Per_Label.Text = "";
                            }

                            T3RBoltStress_SI_Label.Text = Math.Round((ResidualBoltStress / 145.037738007M), 2, MidpointRounding.AwayFromZero).ToString();
                            T3RBoltLoad_In_Label.Text = Math.Round((ResidualBoltStress * TensileStressArea_in / 2240M), 2, MidpointRounding.AwayFromZero).ToString();
                            T3RBoltLoad_SI_Label.Text = Math.Round((ResidualBoltStress * TensileStressArea_mm / 145037.738007M), 2, MidpointRounding.AwayFromZero).ToString();
                            if (BoltYield > 0)
                            {
                                T3RBoltYield_Per_Label.Text = Math.Round(((ResidualBoltStress / 145.037738007M) * 100M / BoltYield), 2, MidpointRounding.AwayFromZero).ToString();
                            }
                            else
                            {
                                T3RBoltYield_Per_Label.Text = "";
                            }
                            DetenBoltStress_In_Label.Text = Math.Round((ResidualBoltStress * LTFValue * (1 + (DetensionProportion / 100M))), 0, MidpointRounding.AwayFromZero).ToString();
                            DetenBoltStress_SI_Label.Text = Math.Round((ResidualBoltStress * LTFValue * (1 + (DetensionProportion / 100M)) / 145.037738007M), 2, MidpointRounding.AwayFromZero).ToString();
                            DetenBoltLoad_In_Label.Text = Math.Round((ResidualBoltStress * LTFValue * (1 + (DetensionProportion / 100M)) * TensileStressArea_in / 2240M), 2, MidpointRounding.AwayFromZero).ToString();
                            DetenBoltLoad_SI_Label.Text = Math.Round((ResidualBoltStress * LTFValue * (1 + (DetensionProportion / 100M)) * TensileStressArea_mm / 145037.738007M), 2, MidpointRounding.AwayFromZero).ToString();
                            if (BoltYield > 0)
                            {
                                DetenBoltYield_Per_Label.Text = Math.Round(((ResidualBoltStress * LTFValue * (1 + (DetensionProportion / 100M)) / 145.037738007M) * 100M / BoltYield), 2, MidpointRounding.AwayFromZero).ToString();
                            }
                            else
                            {
                                DetenBoltYield_Per_Label.Text = "";
                            }

                            // Sequence Tab Calculations for Tool
                            TensioningPressureB = (ResidualBoltStress * TensileStressArea_in * LTFValue / (ToolHydraulicArea / 645.16M));
                            PressureB_In_Label.Text = Math.Round(TensioningPressureB, 0, MidpointRounding.AwayFromZero).ToString() + " psi";
                            PressureB_SI_Label.Text = Math.Round((TensioningPressureB / 145.037738007M), 2, MidpointRounding.AwayFromZero).ToString() + " MPa";
                            PressureA_In_Label.Text = string.Empty;
                            PressureA_SI_Label.Text = string.Empty;
                            Detensioning_In_Label.Text = Math.Round((TensioningPressureB * (1 + (DetensionProportion / 100M))), 0, MidpointRounding.AwayFromZero).ToString() + " psi";
                            Detensioning_SI_Label.Text = Math.Round((TensioningPressureB * (1 + (DetensionProportion / 100M)) / 145.037738007M), 0, MidpointRounding.AwayFromZero).ToString() + " MPa";
                            Pass1AppliedPressure_SequenceTabValue_In_Label.Text = Math.Round(TensioningPressureB, 0, MidpointRounding.AwayFromZero).ToString();
                            Pass1AppliedPressure_SequenceTabValue_SI_Label.Text = Math.Round((TensioningPressureB / 145.037738007M), 2, MidpointRounding.AwayFromZero).ToString();
                            Pass2Bolt_SequenceTabValue_Label.Text = string.Empty;
                            Pass2AppliedPressure_SequenceTabValue_In_Label.Text = string.Empty;
                            Pass2AppliedPressure_SequenceTabValue_SI_Label.Text = string.Empty;
                            Pass3Bolt_SequenceTabValue_Label.Text = string.Empty;
                            Pass3AppliedPressure_SequenceTabValue_In_Label.Text = string.Empty;
                            Pass3AppliedPressure_SequenceTabValue_SI_Label.Text = string.Empty;
                            Pass4Bolt_SequenceTabValue_Label.Text = string.Empty;
                            Pass4AppliedPressure_SequenceTabValue_In_Label.Text = string.Empty;
                            Pass4AppliedPressure_SequenceTabValue_SI_Label.Text = string.Empty;
                            CheckPass1Bolt_SequenceTabValue_Label.Text = "";
                            CheckPass1AppliedPressure_SequenceTabValue_In_Label.Text = "";
                            CheckPass1AppliedPressure_SequenceTabValue_SI_Label.Text = "";
                            CheckPass2Bolt_SequenceTabValue_Label.Text = "";
                            CheckPass2AppliedPressure_SequenceTabValue_In_Label.Text = "";
                            CheckPass2AppliedPressure_SequenceTabValue_SI_Label.Text = "";

                            // For Torque only
                            Torque_SequenceTitle_Label.Visible = false;
                            Torque_SI_Sequence_Label.Visible = false;
                            Torque_In_Sequence_Label.Visible = false;

                            // Warnings that permitted values are exceeded
                            if ((ResidualBoltStress * LTFValue / 145.037738007M) > (BoltYield * 0.95M))     // Only 95% of bolt yield value, pressure A & B allowed
                            {
                                TitleWarning_PictureBox.Visible = true;
                                PressureBWarning_PictureBox.Visible = true;
                            }
                            if ((ResidualBoltStress / 145.037738007M) > (BoltYield * 0.95M))
                            {
                                TitleWarning_PictureBox.Visible = true;
                                ResidualStreeWarning_PictureBox.Visible = true;
                            }
                            if (TensioningPressureB > 21750M)                           // Only 21750 psi tensioning is allowed.
                            {
                                TitleWarning_PictureBox.Visible = true;
                                PressureB_Info_Warning_PictureBox.Visible = true;
                            }
                            if ((TensioningPressureB * (1 + (DetensionProportion / 100M))) > 21750M)
                            {
                                TitleWarning_PictureBox.Visible = true;
                                Deten_Info_Warning_PictureBox.Visible = true;
                            }
                        }
                        else
                        {
                            // Bolt Stress Tab Calculations for Bolt
                            T1BBoltStress_In_Label.Text = "";
                            T1BBoltStress_SI_Label.Text = "";
                            T1BBoltLoad_In_Label.Text = "";
                            T1BBoltLoad_SI_Label.Text = "";
                            T1BBoltYield_Per_Label.Text = "";

                            T1ABoltStress_In_Label.Text = "";
                            T1ABoltStress_SI_Label.Text = "";
                            T1ABoltLoad_In_Label.Text = "";
                            T1ABoltLoad_SI_Label.Text = "";
                            T1ABoltYield_SI_Label.Text = "";

                            T2RBoltStress_In_Label.Text = "";
                            T2RBoltStress_SI_Label.Text = "";
                            T2RBoltLoad_In_Label.Text = "";
                            T2RBoltLoad_SI_Label.Text = "";
                            T2RBoltYield_Per_Label.Text = "";

                            T3RBoltStress_SI_Label.Text = Math.Round((ResidualBoltStress / 145.037738007M), 2, MidpointRounding.AwayFromZero).ToString();
                            T3RBoltLoad_In_Label.Text = Math.Round((ResidualBoltStress * TensileStressArea_in / 2240M), 2, MidpointRounding.AwayFromZero).ToString();
                            T3RBoltLoad_SI_Label.Text = Math.Round((ResidualBoltStress * TensileStressArea_mm / 145037.738007M), 2, MidpointRounding.AwayFromZero).ToString();
                            if (BoltYield > 0)
                            {
                                T3RBoltYield_Per_Label.Text = Math.Round(((ResidualBoltStress / 145.037738007M) * 100M / BoltYield), 2, MidpointRounding.AwayFromZero).ToString();
                            }
                            else
                            {
                                T3RBoltYield_Per_Label.Text = "";
                            }
                            DetenBoltStress_In_Label.Text = "";
                            DetenBoltStress_SI_Label.Text = "";
                            DetenBoltLoad_In_Label.Text = "";
                            DetenBoltLoad_SI_Label.Text = "";
                            DetenBoltYield_Per_Label.Text = "";

                            // Sequence Tab Calculations for Tool
                            PressureB_In_Label.Text = string.Empty;
                            PressureB_SI_Label.Text = string.Empty;
                            PressureA_In_Label.Text = string.Empty;
                            PressureA_SI_Label.Text = string.Empty;
                            Detensioning_In_Label.Text = PressureB_In_Label.Text;
                            Detensioning_SI_Label.Text = PressureB_SI_Label.Text;
                            Pass1AppliedPressure_SequenceTabValue_In_Label.Text = PressureB_In_Label.Text;
                            Pass1AppliedPressure_SequenceTabValue_SI_Label.Text = PressureB_SI_Label.Text;
                            Pass2Bolt_SequenceTabValue_Label.Text = string.Empty;
                            Pass2AppliedPressure_SequenceTabValue_In_Label.Text = string.Empty;
                            Pass2AppliedPressure_SequenceTabValue_SI_Label.Text = string.Empty;
                            Pass3Bolt_SequenceTabValue_Label.Text = string.Empty;
                            Pass3AppliedPressure_SequenceTabValue_In_Label.Text = string.Empty;
                            Pass3AppliedPressure_SequenceTabValue_SI_Label.Text = string.Empty;
                            Pass4Bolt_SequenceTabValue_Label.Text = string.Empty;
                            Pass4AppliedPressure_SequenceTabValue_In_Label.Text = string.Empty;
                            Pass4AppliedPressure_SequenceTabValue_SI_Label.Text = string.Empty;
                            CheckPass1Bolt_SequenceTabValue_Label.Text = "";
                            CheckPass1AppliedPressure_SequenceTabValue_In_Label.Text = "";
                            CheckPass1AppliedPressure_SequenceTabValue_SI_Label.Text = "";
                            CheckPass2Bolt_SequenceTabValue_Label.Text = "";
                            CheckPass2AppliedPressure_SequenceTabValue_In_Label.Text = "";
                            CheckPass2AppliedPressure_SequenceTabValue_SI_Label.Text = "";
                            // For Torque only
                            Torque_SequenceTitle_Label.Visible = false;
                            Torque_SI_Sequence_Label.Visible = false;
                            Torque_In_Sequence_Label.Visible = false;
                        }
                    }
                    else
                    {
                        if (ToolHydraulicArea > 0)
                        {
                            // Bolt Stress Tab Calculations for Bolt
                            T1BBoltStress_In_Label.Text = Math.Round((ResidualBoltStress * 145.037738007M * LTFValue), 0, MidpointRounding.AwayFromZero).ToString();
                            T1BBoltStress_SI_Label.Text = Math.Round((ResidualBoltStress * LTFValue), 2, MidpointRounding.AwayFromZero).ToString();
                            T1BBoltLoad_In_Label.Text = Math.Round((ResidualBoltStress * 145.037738007M * LTFValue * TensileStressArea_in / 2240M), 2, MidpointRounding.AwayFromZero).ToString();
                            T1BBoltLoad_SI_Label.Text = Math.Round((ResidualBoltStress * TensileStressArea_mm * LTFValue / 1000M), 2, MidpointRounding.AwayFromZero).ToString();
                            if (BoltYield > 0)
                            {
                                T1BBoltYield_Per_Label.Text = Math.Round(((ResidualBoltStress * LTFValue) * 100M / BoltYield), 2, MidpointRounding.AwayFromZero).ToString();
                            }
                            else
                            {
                                T1BBoltYield_Per_Label.Text = "";
                            }

                            T1ABoltStress_In_Label.Text = "";
                            T1ABoltStress_SI_Label.Text = "";
                            T1ABoltLoad_In_Label.Text = "";
                            T1ABoltLoad_SI_Label.Text = "";
                            T1ABoltYield_SI_Label.Text = "";

                            T2RBoltStress_In_Label.Text = Math.Round((ResidualBoltStress * 145.037738007M), 0, MidpointRounding.AwayFromZero).ToString();
                            T2RBoltStress_SI_Label.Text = ResidualBoltStress_TextBox.Text;
                            T2RBoltLoad_In_Label.Text = Math.Round((ResidualBoltStress * 145.037738007M * TensileStressArea_in / 2240M), 2, MidpointRounding.AwayFromZero).ToString();
                            T2RBoltLoad_SI_Label.Text = Math.Round((ResidualBoltStress * TensileStressArea_mm / 1000M), 2, MidpointRounding.AwayFromZero).ToString();
                            if (BoltYield > 0)
                            {
                                T2RBoltYield_Per_Label.Text = Math.Round(((ResidualBoltStress) * 100M / BoltYield), 2, MidpointRounding.AwayFromZero).ToString();
                            }
                            else
                            {
                                T2RBoltYield_Per_Label.Text = "";
                            }
                            T3RBoltStress_In_Label.Text = Math.Round((ResidualBoltStress * 145.037738007M), 0, MidpointRounding.AwayFromZero).ToString();
                            T3RBoltStress_SI_Label.Text = ResidualBoltStress_TextBox.Text;
                            T3RBoltLoad_In_Label.Text = Math.Round((ResidualBoltStress * 145.037738007M * TensileStressArea_in / 2240M), 2, MidpointRounding.AwayFromZero).ToString();
                            T3RBoltLoad_SI_Label.Text = Math.Round((ResidualBoltStress * TensileStressArea_mm / 1000M), 2, MidpointRounding.AwayFromZero).ToString();
                            if (BoltYield > 0)
                            {
                                T3RBoltYield_Per_Label.Text = Math.Round(((ResidualBoltStress) * 100M / BoltYield), 2, MidpointRounding.AwayFromZero).ToString();
                            }
                            else
                            {
                                T3RBoltYield_Per_Label.Text = "";
                            }
                            DetenBoltStress_In_Label.Text = Math.Round((ResidualBoltStress * 145.037738007M * LTFValue * (1 + (DetensionProportion / 100M))), 0, MidpointRounding.AwayFromZero).ToString();
                            DetenBoltStress_SI_Label.Text = Math.Round((ResidualBoltStress * LTFValue * (1 + (DetensionProportion / 100M))), 2, MidpointRounding.AwayFromZero).ToString();
                            DetenBoltLoad_In_Label.Text = Math.Round((ResidualBoltStress * 145.037738007M * LTFValue * (1 + (DetensionProportion / 100M)) * TensileStressArea_in / 2240M), 2, MidpointRounding.AwayFromZero).ToString();
                            DetenBoltLoad_SI_Label.Text = Math.Round((ResidualBoltStress * LTFValue * (1 + (DetensionProportion / 100M)) * TensileStressArea_mm / 1000M), 2, MidpointRounding.AwayFromZero).ToString();
                            if (BoltYield > 0)
                            {
                                DetenBoltYield_Per_Label.Text = Math.Round(((ResidualBoltStress * LTFValue) * (1 + (DetensionProportion / 100M)) * 100M / BoltYield), 2, MidpointRounding.AwayFromZero).ToString();
                            }
                            else
                            {
                                DetenBoltYield_Per_Label.Text = "";
                            }
                            // Sequence Tab Calculations for Tool
                            if (ToolHydraulicArea > 0)
                            {
                                TensioningPressureB = (ResidualBoltStress * TensileStressArea_mm * LTFValue / ToolHydraulicArea);
                            }
                            PressureB_In_Label.Text = Math.Round((TensioningPressureB * 145.037738007M), 0, MidpointRounding.AwayFromZero).ToString() + " psi";
                            PressureB_SI_Label.Text = Math.Round((TensioningPressureB), 2, MidpointRounding.AwayFromZero).ToString() + " MPa";
                            PressureA_In_Label.Text = string.Empty;
                            PressureA_SI_Label.Text = string.Empty;
                            Detensioning_In_Label.Text = Math.Round((TensioningPressureB * (1 + (DetensionProportion / 100M)) * 145.037738007M), 0, MidpointRounding.AwayFromZero).ToString() + " psi";
                            Detensioning_SI_Label.Text = Math.Round((TensioningPressureB * (1 + (DetensionProportion / 100M))), 0, MidpointRounding.AwayFromZero).ToString() + " MPa";
                            Pass1AppliedPressure_SequenceTabValue_In_Label.Text = Math.Round((TensioningPressureB * 145.037738007M), 0, MidpointRounding.AwayFromZero).ToString();
                            Pass1AppliedPressure_SequenceTabValue_SI_Label.Text = Math.Round((TensioningPressureB), 2, MidpointRounding.AwayFromZero).ToString();
                            Pass2Bolt_SequenceTabValue_Label.Text = string.Empty;
                            Pass2AppliedPressure_SequenceTabValue_In_Label.Text = string.Empty;
                            Pass2AppliedPressure_SequenceTabValue_SI_Label.Text = string.Empty;
                            Pass3Bolt_SequenceTabValue_Label.Text = string.Empty;
                            Pass3AppliedPressure_SequenceTabValue_In_Label.Text = string.Empty;
                            Pass3AppliedPressure_SequenceTabValue_SI_Label.Text = string.Empty;
                            Pass4Bolt_SequenceTabValue_Label.Text = string.Empty;
                            Pass4AppliedPressure_SequenceTabValue_In_Label.Text = string.Empty;
                            Pass4AppliedPressure_SequenceTabValue_SI_Label.Text = string.Empty;
                            CheckPass1Bolt_SequenceTabValue_Label.Text = "";
                            CheckPass1AppliedPressure_SequenceTabValue_In_Label.Text = "";
                            CheckPass1AppliedPressure_SequenceTabValue_SI_Label.Text = "";
                            CheckPass2Bolt_SequenceTabValue_Label.Text = "";
                            CheckPass2AppliedPressure_SequenceTabValue_In_Label.Text = "";
                            CheckPass2AppliedPressure_SequenceTabValue_SI_Label.Text = "";
                            // For Torque only
                            Torque_SequenceTitle_Label.Visible = false;
                            Torque_SI_Sequence_Label.Visible = false;
                            Torque_In_Sequence_Label.Visible = false;
                            // Warnings that permitted values are exceeded
                            if ((ResidualBoltStress * LTFValue) > (BoltYield * 0.95M))     // Only 95% of bolt yield value, pressure A & B allowed
                            {
                                TitleWarning_PictureBox.Visible = true;
                                PressureBWarning_PictureBox.Visible = true;
                            }
                            if ((ResidualBoltStress) > (BoltYield * 0.95M))
                            {
                                TitleWarning_PictureBox.Visible = true;
                                ResidualStreeWarning_PictureBox.Visible = true;
                            }
                            if (TensioningPressureB > 150)                           // Only 150 MPa tensioning is allowed.
                            {
                                TitleWarning_PictureBox.Visible = true;
                                PressureB_Info_Warning_PictureBox.Visible = true;
                            }
                            if ((TensioningPressureB * (1 + (DetensionProportion / 100M))) > 150)
                            {
                                TitleWarning_PictureBox.Visible = true;
                                Deten_Info_Warning_PictureBox.Visible = true;
                            }
                        }
                        else
                        {
                            // Bolt Stress Tab Calculations for Bolt
                            T1BBoltStress_In_Label.Text = "";
                            T1BBoltStress_SI_Label.Text = "";
                            T1BBoltLoad_In_Label.Text = "";
                            T1BBoltLoad_SI_Label.Text = "";
                            T1BBoltYield_Per_Label.Text = "";

                            T1ABoltStress_In_Label.Text = "";
                            T1ABoltStress_SI_Label.Text = "";
                            T1ABoltLoad_In_Label.Text = "";
                            T1ABoltLoad_SI_Label.Text = "";
                            T1ABoltYield_SI_Label.Text = "";

                            T2RBoltStress_In_Label.Text = "";
                            T2RBoltStress_SI_Label.Text = "";
                            T2RBoltLoad_In_Label.Text = "";
                            T2RBoltLoad_SI_Label.Text = "";
                            T2RBoltYield_Per_Label.Text = "";

                            T3RBoltStress_In_Label.Text = Math.Round((ResidualBoltStress * 145.037738007M), 0, MidpointRounding.AwayFromZero).ToString();
                            T3RBoltStress_SI_Label.Text = ResidualBoltStress_TextBox.Text;
                            T3RBoltLoad_In_Label.Text = Math.Round((ResidualBoltStress * 145.037738007M * TensileStressArea_in / 2240M), 2, MidpointRounding.AwayFromZero).ToString();
                            T3RBoltLoad_SI_Label.Text = Math.Round((ResidualBoltStress * TensileStressArea_mm / 1000M), 2, MidpointRounding.AwayFromZero).ToString();
                            if (BoltYield > 0)
                            {
                                T3RBoltYield_Per_Label.Text = Math.Round(((ResidualBoltStress) * 100M / BoltYield), 2, MidpointRounding.AwayFromZero).ToString();
                            }
                            else
                            {
                                T3RBoltYield_Per_Label.Text = "";
                            }
                            DetenBoltStress_In_Label.Text = "";
                            DetenBoltStress_SI_Label.Text = "";
                            DetenBoltLoad_In_Label.Text = "";
                            DetenBoltLoad_SI_Label.Text = "";
                            DetenBoltYield_Per_Label.Text = "";

                            // Sequence Tab Calculations for Tool
                            PressureB_In_Label.Text = string.Empty;
                            PressureB_SI_Label.Text = string.Empty;
                            PressureA_In_Label.Text = string.Empty;
                            PressureA_SI_Label.Text = string.Empty;
                            Detensioning_In_Label.Text = PressureB_In_Label.Text;
                            Detensioning_SI_Label.Text = PressureB_SI_Label.Text;
                            Pass1AppliedPressure_SequenceTabValue_In_Label.Text = PressureB_In_Label.Text;
                            Pass1AppliedPressure_SequenceTabValue_SI_Label.Text = PressureB_SI_Label.Text;
                            Pass2Bolt_SequenceTabValue_Label.Text = string.Empty;
                            Pass2AppliedPressure_SequenceTabValue_In_Label.Text = string.Empty;
                            Pass2AppliedPressure_SequenceTabValue_SI_Label.Text = string.Empty;
                            Pass3Bolt_SequenceTabValue_Label.Text = string.Empty;
                            Pass3AppliedPressure_SequenceTabValue_In_Label.Text = string.Empty;
                            Pass3AppliedPressure_SequenceTabValue_SI_Label.Text = string.Empty;
                            Pass4Bolt_SequenceTabValue_Label.Text = string.Empty;
                            Pass4AppliedPressure_SequenceTabValue_In_Label.Text = string.Empty;
                            Pass4AppliedPressure_SequenceTabValue_SI_Label.Text = string.Empty;
                            CheckPass1Bolt_SequenceTabValue_Label.Text = "";
                            CheckPass1AppliedPressure_SequenceTabValue_In_Label.Text = "";
                            CheckPass1AppliedPressure_SequenceTabValue_SI_Label.Text = "";
                            CheckPass2Bolt_SequenceTabValue_Label.Text = "";
                            CheckPass2AppliedPressure_SequenceTabValue_In_Label.Text = "";
                            CheckPass2AppliedPressure_SequenceTabValue_SI_Label.Text = "";
                            // For Torque only
                            Torque_SequenceTitle_Label.Visible = false;
                            Torque_SI_Sequence_Label.Visible = false;
                            Torque_In_Sequence_Label.Visible = false;
                        }
                    }
                    if (NumberOfBolts_TextBox.Text.Trim() != "")
                    {
                        int Bolts = Convert.ToInt32(NumberOfBolts_TextBox.Text);
                        switch (Bolts)
                        {
                            case 4:
                                Sequence_PictureBox.Image = WizBolt.Properties.Resources.Flange100_04;
                                break;
                            case 8:
                                Sequence_PictureBox.Image = WizBolt.Properties.Resources.Flange100_08;
                                break;
                            case 12:
                                Sequence_PictureBox.Image = WizBolt.Properties.Resources.Flange100_12;
                                break;
                            case 16:
                                Sequence_PictureBox.Image = WizBolt.Properties.Resources.Flange100_16;
                                break;
                            case 20:
                                Sequence_PictureBox.Image = WizBolt.Properties.Resources.Flange100_20;
                                break;
                            case 24:
                                Sequence_PictureBox.Image = WizBolt.Properties.Resources.Flange100_24;
                                break;
                            case 28:
                                Sequence_PictureBox.Image = WizBolt.Properties.Resources.Flange100_28;
                                break;
                            case 32:
                                Sequence_PictureBox.Image = WizBolt.Properties.Resources.Flange100_32;
                                break;
                            case 36:
                                Sequence_PictureBox.Image = WizBolt.Properties.Resources.Flange100_36;
                                break;
                            case 40:
                                Sequence_PictureBox.Image = WizBolt.Properties.Resources.Flange100_40;
                                break;
                            case 44:
                                Sequence_PictureBox.Image = WizBolt.Properties.Resources.Flange100_44;
                                break;
                            case 48:
                                Sequence_PictureBox.Image = WizBolt.Properties.Resources.Flange100_48;
                                break;
                            case 52:
                                Sequence_PictureBox.Image = WizBolt.Properties.Resources.Flange100_52;
                                break;
                            case 56:
                                Sequence_PictureBox.Image = WizBolt.Properties.Resources.Flange100_56;
                                break;
                            case 60:
                                Sequence_PictureBox.Image = WizBolt.Properties.Resources.Flange100_60;
                                break;
                            default:
                                Sequence_PictureBox.Image = WizBolt.Properties.Resources.Sequence_Full;
                                break;
                        }
                    }
                    else
                    {
                        Sequence_PictureBox.Image = null;
                    }
                    break;
                case 1:
                    Sequence_PictureBox.Image = WizBolt.Properties.Resources.Sequence_half;
                    if (UnitSystem_Button.Text == "SAE")
                    {
                        if (ToolHydraulicArea > 0)
                        {
                            // Bolt Stress Tab Calculations for Bolt
                            T1BBoltStress_In_Label.Text = Math.Round((ResidualBoltStress * LTFValue), 0, MidpointRounding.AwayFromZero).ToString();
                            T1BBoltStress_SI_Label.Text = Math.Round((ResidualBoltStress * LTFValue / 145.037738007M), 2, MidpointRounding.AwayFromZero).ToString();
                            T1BBoltLoad_In_Label.Text = Math.Round((ResidualBoltStress * LTFValue * TensileStressArea_in / 2240M), 2, MidpointRounding.AwayFromZero).ToString();
                            T1BBoltLoad_SI_Label.Text = Math.Round((ResidualBoltStress * TensileStressArea_mm * LTFValue / 145037.738007M), 2, MidpointRounding.AwayFromZero).ToString();
                            if (BoltYield > 0)
                            {
                                T1BBoltYield_Per_Label.Text = Math.Round(((ResidualBoltStress * LTFValue / 145.037738007M) * 100M / BoltYield), 2, MidpointRounding.AwayFromZero).ToString();
                            }
                            else
                            {
                                T1BBoltYield_Per_Label.Text = "";
                            }
                            T1ABoltStress_In_Label.Text = Math.Round((ResidualBoltStress * LTFValue * (1 + (CrossLoad / 100M))), 0, MidpointRounding.AwayFromZero).ToString(); ;
                            T1ABoltStress_SI_Label.Text = Math.Round((ResidualBoltStress * LTFValue * (1 + (CrossLoad / 100M)) / 145.037738007M), 2, MidpointRounding.AwayFromZero).ToString(); ;
                            T1ABoltLoad_In_Label.Text = Math.Round((ResidualBoltStress * LTFValue * (1 + (CrossLoad / 100M)) * TensileStressArea_in / 2240M), 2, MidpointRounding.AwayFromZero).ToString();
                            T1ABoltLoad_SI_Label.Text = Math.Round((ResidualBoltStress * (1 + (CrossLoad / 100M)) * TensileStressArea_mm * LTFValue / 145037.738007M), 2, MidpointRounding.AwayFromZero).ToString();
                            if (BoltYield > 0)
                            {
                                T1ABoltYield_SI_Label.Text = Math.Round(((ResidualBoltStress * (1 + (CrossLoad / 100M)) * LTFValue / 145.037738007M) * 100M / BoltYield), 2, MidpointRounding.AwayFromZero).ToString();
                            }
                            else
                            {
                                T1ABoltYield_SI_Label.Text = "";
                            }

                            T2RBoltStress_In_Label.Text = ResidualBoltStress_TextBox.Text;
                            T2RBoltStress_SI_Label.Text = Math.Round((ResidualBoltStress / 145.037738007M), 2, MidpointRounding.AwayFromZero).ToString();
                            T2RBoltLoad_In_Label.Text = Math.Round((ResidualBoltStress * TensileStressArea_in / 2240M), 2, MidpointRounding.AwayFromZero).ToString();
                            T2RBoltLoad_SI_Label.Text = Math.Round((ResidualBoltStress * TensileStressArea_mm / 145037.738007M), 2, MidpointRounding.AwayFromZero).ToString();
                            if (BoltYield > 0)
                            {
                                T2RBoltYield_Per_Label.Text = Math.Round(((ResidualBoltStress / 145.037738007M) * 100M / BoltYield), 2, MidpointRounding.AwayFromZero).ToString();
                            }
                            else
                            {
                                T2RBoltYield_Per_Label.Text = "";
                            }

                            T3RBoltStress_In_Label.Text = ResidualBoltStress_TextBox.Text;
                            T3RBoltStress_SI_Label.Text = Math.Round((ResidualBoltStress / 145.037738007M), 2, MidpointRounding.AwayFromZero).ToString();
                            T3RBoltLoad_In_Label.Text = Math.Round((ResidualBoltStress * TensileStressArea_in / 2240M), 2, MidpointRounding.AwayFromZero).ToString();
                            T3RBoltLoad_SI_Label.Text = Math.Round((ResidualBoltStress * TensileStressArea_mm / 145037.738007M), 2, MidpointRounding.AwayFromZero).ToString();
                            if (BoltYield > 0)
                            {
                                T3RBoltYield_Per_Label.Text = Math.Round(((ResidualBoltStress / 145.037738007M) * 100M / BoltYield), 2, MidpointRounding.AwayFromZero).ToString();
                            }
                            else
                            {
                                T3RBoltYield_Per_Label.Text = "";
                            }
                            DetenBoltStress_In_Label.Text = Math.Round((ResidualBoltStress * LTFValue * (1 + (DetensionProportion / 100M))), 0, MidpointRounding.AwayFromZero).ToString();
                            DetenBoltStress_SI_Label.Text = Math.Round((ResidualBoltStress * LTFValue * (1 + (DetensionProportion / 100M)) / 145.037738007M), 2, MidpointRounding.AwayFromZero).ToString();
                            DetenBoltLoad_In_Label.Text = Math.Round((ResidualBoltStress * LTFValue * (1 + (DetensionProportion / 100M)) * TensileStressArea_in / 2240M), 2, MidpointRounding.AwayFromZero).ToString();
                            DetenBoltLoad_SI_Label.Text = Math.Round((ResidualBoltStress * LTFValue * (1 + (DetensionProportion / 100M)) * TensileStressArea_mm / 145037.738007M), 2, MidpointRounding.AwayFromZero).ToString();
                            if (BoltYield > 0)
                            {
                                DetenBoltYield_Per_Label.Text = Math.Round(((ResidualBoltStress * LTFValue * (1 + (DetensionProportion / 100M)) / 145.037738007M) * 100M / BoltYield), 2, MidpointRounding.AwayFromZero).ToString();
                            }
                            else
                            {
                                DetenBoltYield_Per_Label.Text = "";
                            }

                            // Sequence Tab Calculations for Tool

                            TensioningPressureB = (ResidualBoltStress * TensileStressArea_in * LTFValue / (ToolHydraulicArea / 645.16M));
                            TensioningPressureA = TensioningPressureB * (1 + (CrossLoad / 100));
                            PressureB_In_Label.Text = Math.Round(TensioningPressureB, 0, MidpointRounding.AwayFromZero).ToString() + " psi";
                            PressureB_SI_Label.Text = Math.Round((TensioningPressureB / 145.037738007M), 2, MidpointRounding.AwayFromZero).ToString() + " MPa";
                            Detensioning_In_Label.Text = Math.Round((TensioningPressureB * (1 + (DetensionProportion / 100M))), 0, MidpointRounding.AwayFromZero).ToString() + " psi";
                            Detensioning_SI_Label.Text = Math.Round((TensioningPressureB * (1 + (DetensionProportion / 100M)) / 145.037738007M), 0, MidpointRounding.AwayFromZero).ToString() + " MPa";
                            PressureA_In_Label.Text = Math.Round(TensioningPressureA, 0, MidpointRounding.AwayFromZero).ToString() + " psi";
                            PressureA_SI_Label.Text = Math.Round((TensioningPressureA / 145.037738007M), 2, MidpointRounding.AwayFromZero).ToString() + " MPa";
                            Pass1AppliedPressure_SequenceTabValue_In_Label.Text = Math.Round(TensioningPressureA, 0, MidpointRounding.AwayFromZero).ToString();
                            Pass1AppliedPressure_SequenceTabValue_SI_Label.Text = Math.Round((TensioningPressureA / 145.037738007M), 2, MidpointRounding.AwayFromZero).ToString();
                            Pass2Bolt_SequenceTabValue_Label.Text = "2";
                            Pass2AppliedPressure_SequenceTabValue_In_Label.Text = Math.Round(TensioningPressureB, 0, MidpointRounding.AwayFromZero).ToString();
                            Pass2AppliedPressure_SequenceTabValue_SI_Label.Text = Math.Round((TensioningPressureB / 145.037738007M), 2, MidpointRounding.AwayFromZero).ToString();
                            Pass3Bolt_SequenceTabValue_Label.Text = string.Empty;
                            Pass3AppliedPressure_SequenceTabValue_In_Label.Text = string.Empty;
                            Pass3AppliedPressure_SequenceTabValue_SI_Label.Text = string.Empty;
                            Pass4Bolt_SequenceTabValue_Label.Text = string.Empty;
                            Pass4AppliedPressure_SequenceTabValue_In_Label.Text = string.Empty;
                            Pass4AppliedPressure_SequenceTabValue_SI_Label.Text = string.Empty;
                            CheckPass1Bolt_SequenceTabValue_Label.Text = "1";
                            CheckPass1AppliedPressure_SequenceTabValue_In_Label.Text = Math.Round(TensioningPressureB, 0, MidpointRounding.AwayFromZero).ToString();
                            CheckPass1AppliedPressure_SequenceTabValue_SI_Label.Text = Math.Round((TensioningPressureB / 145.037738007M), 2, MidpointRounding.AwayFromZero).ToString();
                            CheckPass2Bolt_SequenceTabValue_Label.Text = "2";
                            CheckPass2AppliedPressure_SequenceTabValue_In_Label.Text = Math.Round(TensioningPressureB, 0, MidpointRounding.AwayFromZero).ToString();
                            CheckPass2AppliedPressure_SequenceTabValue_SI_Label.Text = Math.Round((TensioningPressureB / 145.037738007M), 2, MidpointRounding.AwayFromZero).ToString();
                            // For Torque only
                            Torque_SequenceTitle_Label.Visible = false;
                            Torque_SI_Sequence_Label.Visible = false;
                            Torque_In_Sequence_Label.Visible = false;
                            // Warnings that permitted values are exceeded
                            if ((ResidualBoltStress * LTFValue * (1 + (CrossLoad / 100M)) / 145.037738007M) > (BoltYield * 0.95M))     // Only 95% of bolt yield value, pressure A & B allowed
                            {
                                TitleWarning_PictureBox.Visible = true;
                                PressureAWarning_PictureBox.Visible = true;
                            }
                            if ((ResidualBoltStress * LTFValue / 145.037738007M) > (BoltYield * 0.95M))     // Only 95% of bolt yield value, pressure A & B allowed
                            {
                                TitleWarning_PictureBox.Visible = true;
                                PressureBWarning_PictureBox.Visible = true;
                            }
                            if ((ResidualBoltStress / 145.037738007M) > (BoltYield * 0.95M))
                            {
                                TitleWarning_PictureBox.Visible = true;
                                ResidualStreeWarning_PictureBox.Visible = true;
                            }
                            if (TensioningPressureA > 21750M)                           // Only 21750 psi tensioning is allowed.
                            {
                                TitleWarning_PictureBox.Visible = true;
                                PressureA_Info_Warning_PictureBox.Visible = true;
                            }
                            if (TensioningPressureB > 21750M)                           // Only 21750 psi tensioning is allowed.
                            {
                                TitleWarning_PictureBox.Visible = true;
                                PressureB_Info_Warning_PictureBox.Visible = true;
                            }
                            if ((TensioningPressureB * (1 + (DetensionProportion / 100M))) > 21750M)
                            {
                                TitleWarning_PictureBox.Visible = true;
                                Deten_Info_Warning_PictureBox.Visible = true;
                            }

                        }
                        else
                        {
                            // Bolt Stress Tab Calculations for Bolt
                            T1BBoltStress_In_Label.Text = "";
                            T1BBoltStress_SI_Label.Text = "";
                            T1BBoltLoad_In_Label.Text = "";
                            T1BBoltLoad_SI_Label.Text = "";
                            T1BBoltYield_Per_Label.Text = "";

                            T1ABoltStress_In_Label.Text = "";
                            T1ABoltStress_SI_Label.Text = "";
                            T1ABoltLoad_In_Label.Text = "";
                            T1ABoltLoad_SI_Label.Text = "";
                            T1ABoltYield_SI_Label.Text = "";

                            T2RBoltStress_In_Label.Text = "";
                            T2RBoltStress_SI_Label.Text = "";
                            T2RBoltLoad_In_Label.Text = "";
                            T2RBoltLoad_SI_Label.Text = "";
                            T2RBoltYield_Per_Label.Text = "";

                            T3RBoltStress_In_Label.Text = ResidualBoltStress_TextBox.Text;
                            T3RBoltStress_SI_Label.Text = Math.Round((ResidualBoltStress / 145.037738007M), 2, MidpointRounding.AwayFromZero).ToString();
                            T3RBoltLoad_In_Label.Text = Math.Round((ResidualBoltStress * TensileStressArea_in / 2240M), 2, MidpointRounding.AwayFromZero).ToString();
                            T3RBoltLoad_SI_Label.Text = Math.Round((ResidualBoltStress * TensileStressArea_mm / 145037.738007M), 2, MidpointRounding.AwayFromZero).ToString();
                            if (BoltYield > 0)
                            {
                                T3RBoltYield_Per_Label.Text = Math.Round(((ResidualBoltStress / 145.037738007M) * 100M / BoltYield), 2, MidpointRounding.AwayFromZero).ToString();
                            }
                            else
                            {
                                T3RBoltYield_Per_Label.Text = "";
                            }
                            DetenBoltStress_In_Label.Text = "";
                            DetenBoltStress_SI_Label.Text = "";
                            DetenBoltLoad_In_Label.Text = "";
                            DetenBoltLoad_SI_Label.Text = "";
                            DetenBoltYield_Per_Label.Text = "";

                            // Sequence Tab Calculations for Tool
                            PressureB_In_Label.Text = "";
                            PressureB_SI_Label.Text = "";
                            Detensioning_In_Label.Text = PressureB_In_Label.Text;
                            Detensioning_SI_Label.Text = PressureB_SI_Label.Text;
                            PressureA_In_Label.Text = string.Empty;
                            PressureA_SI_Label.Text = string.Empty;
                            Pass1AppliedPressure_SequenceTabValue_In_Label.Text = PressureA_In_Label.Text;
                            Pass1AppliedPressure_SequenceTabValue_SI_Label.Text = PressureA_SI_Label.Text;
                            Pass2Bolt_SequenceTabValue_Label.Text = "";
                            Pass2AppliedPressure_SequenceTabValue_In_Label.Text = PressureB_In_Label.Text;
                            Pass2AppliedPressure_SequenceTabValue_SI_Label.Text = PressureB_SI_Label.Text;
                            Pass3Bolt_SequenceTabValue_Label.Text = string.Empty;
                            Pass3AppliedPressure_SequenceTabValue_In_Label.Text = string.Empty;
                            Pass3AppliedPressure_SequenceTabValue_SI_Label.Text = string.Empty;
                            Pass4Bolt_SequenceTabValue_Label.Text = string.Empty;
                            Pass4AppliedPressure_SequenceTabValue_In_Label.Text = string.Empty;
                            Pass4AppliedPressure_SequenceTabValue_SI_Label.Text = string.Empty;
                            CheckPass1Bolt_SequenceTabValue_Label.Text = "";
                            CheckPass1AppliedPressure_SequenceTabValue_In_Label.Text = PressureB_In_Label.Text;
                            CheckPass1AppliedPressure_SequenceTabValue_SI_Label.Text = PressureB_SI_Label.Text;
                            CheckPass2Bolt_SequenceTabValue_Label.Text = "";
                            CheckPass2AppliedPressure_SequenceTabValue_In_Label.Text = PressureB_In_Label.Text;
                            CheckPass2AppliedPressure_SequenceTabValue_SI_Label.Text = PressureB_SI_Label.Text;
                            // For Torque only
                            Torque_SequenceTitle_Label.Visible = false;
                            Torque_SI_Sequence_Label.Visible = false;
                            Torque_In_Sequence_Label.Visible = false;
                        }
                    }
                    else
                    {
                        if (ToolHydraulicArea > 0)
                        {
                            // Bolt Stress Tab Calculations for Bolt
                            T1BBoltStress_In_Label.Text = Math.Round((ResidualBoltStress * 145.037738007M * LTFValue), 0, MidpointRounding.AwayFromZero).ToString();
                            T1BBoltStress_SI_Label.Text = Math.Round((ResidualBoltStress * LTFValue), 2, MidpointRounding.AwayFromZero).ToString();
                            T1BBoltLoad_In_Label.Text = Math.Round((ResidualBoltStress * 145.037738007M * LTFValue * TensileStressArea_in / 2240M), 2, MidpointRounding.AwayFromZero).ToString();
                            T1BBoltLoad_SI_Label.Text = Math.Round((ResidualBoltStress * TensileStressArea_mm * LTFValue / 1000M), 2, MidpointRounding.AwayFromZero).ToString();
                            if (BoltYield > 0)
                            {
                                T1BBoltYield_Per_Label.Text = Math.Round(((ResidualBoltStress * LTFValue) * 100M / BoltYield), 2, MidpointRounding.AwayFromZero).ToString();
                            }
                            else
                            {
                                T1BBoltYield_Per_Label.Text = "";
                            }

                            T1ABoltStress_In_Label.Text = Math.Round((ResidualBoltStress * 145.037738007M * LTFValue * (1 + (CrossLoad / 100M))), 0, MidpointRounding.AwayFromZero).ToString();
                            T1ABoltStress_SI_Label.Text = Math.Round((ResidualBoltStress * LTFValue * (1 + (CrossLoad / 100M))), 2, MidpointRounding.AwayFromZero).ToString();
                            T1ABoltLoad_In_Label.Text = Math.Round((ResidualBoltStress * 145.037738007M * LTFValue * (1 + (CrossLoad / 100M)) * TensileStressArea_in / 2240M), 2, MidpointRounding.AwayFromZero).ToString();
                            T1ABoltLoad_SI_Label.Text = Math.Round((ResidualBoltStress * TensileStressArea_mm * LTFValue * (1 + (CrossLoad / 100M)) / 1000M), 2, MidpointRounding.AwayFromZero).ToString();
                            if (BoltYield > 0)
                            {
                                T1ABoltYield_SI_Label.Text = Math.Round(((ResidualBoltStress * LTFValue * (1 + (CrossLoad / 100M))) * 100M / BoltYield), 2, MidpointRounding.AwayFromZero).ToString();
                            }
                            else
                            {
                                T1ABoltYield_SI_Label.Text = "";
                            }

                            T2RBoltStress_In_Label.Text = Math.Round((ResidualBoltStress * 145.037738007M), 0, MidpointRounding.AwayFromZero).ToString();
                            T2RBoltStress_SI_Label.Text = ResidualBoltStress_TextBox.Text;
                            T2RBoltLoad_In_Label.Text = Math.Round((ResidualBoltStress * 145.037738007M * TensileStressArea_in / 2240M), 2, MidpointRounding.AwayFromZero).ToString();
                            T2RBoltLoad_SI_Label.Text = Math.Round((ResidualBoltStress * TensileStressArea_mm / 1000M), 2, MidpointRounding.AwayFromZero).ToString();
                            if (BoltYield > 0)
                            {
                                T2RBoltYield_Per_Label.Text = Math.Round(((ResidualBoltStress) * 100M / BoltYield), 2, MidpointRounding.AwayFromZero).ToString();
                            }
                            else
                            {
                                T2RBoltYield_Per_Label.Text = "";
                            }
                            T3RBoltStress_In_Label.Text = Math.Round((ResidualBoltStress * 145.037738007M), 0, MidpointRounding.AwayFromZero).ToString();
                            T3RBoltStress_SI_Label.Text = ResidualBoltStress_TextBox.Text;
                            T3RBoltLoad_In_Label.Text = Math.Round((ResidualBoltStress * 145.037738007M * TensileStressArea_in / 2240M), 2, MidpointRounding.AwayFromZero).ToString();
                            T3RBoltLoad_SI_Label.Text = Math.Round((ResidualBoltStress * TensileStressArea_mm / 1000M), 2, MidpointRounding.AwayFromZero).ToString();
                            if (BoltYield > 0)
                            {
                                T3RBoltYield_Per_Label.Text = Math.Round(((ResidualBoltStress) * 100M / BoltYield), 2, MidpointRounding.AwayFromZero).ToString();
                            }
                            else
                            {
                                T3RBoltYield_Per_Label.Text = "";
                            }
                            DetenBoltStress_In_Label.Text = Math.Round((ResidualBoltStress * 145.037738007M * LTFValue * (1 + (DetensionProportion / 100M))), 0, MidpointRounding.AwayFromZero).ToString();
                            DetenBoltStress_SI_Label.Text = Math.Round((ResidualBoltStress * LTFValue * (1 + (DetensionProportion / 100M))), 2, MidpointRounding.AwayFromZero).ToString();
                            DetenBoltLoad_In_Label.Text = Math.Round((ResidualBoltStress * 145.037738007M * LTFValue * (1 + (DetensionProportion / 100M)) * TensileStressArea_in / 2240M), 2, MidpointRounding.AwayFromZero).ToString();
                            DetenBoltLoad_SI_Label.Text = Math.Round((ResidualBoltStress * LTFValue * (1 + (DetensionProportion / 100M)) * TensileStressArea_mm / 1000M), 2, MidpointRounding.AwayFromZero).ToString();
                            if (BoltYield > 0)
                            {
                                DetenBoltYield_Per_Label.Text = Math.Round(((ResidualBoltStress * LTFValue * (1 + (DetensionProportion / 100M))) * 100M / BoltYield), 2, MidpointRounding.AwayFromZero).ToString();
                            }
                            else
                            {
                                DetenBoltYield_Per_Label.Text = "";
                            }

                            // Sequence Tab Calculations for Tool
                            if (ToolHydraulicArea > 0)
                            {
                                TensioningPressureB = (ResidualBoltStress * TensileStressArea_mm * LTFValue / ToolHydraulicArea);
                            }
                            TensioningPressureA = TensioningPressureB * (1 + (CrossLoad / 100));
                            PressureB_In_Label.Text = Math.Round((TensioningPressureB * 145.037738007M), 0, MidpointRounding.AwayFromZero).ToString() + " psi";
                            PressureB_SI_Label.Text = Math.Round((TensioningPressureB), 2, MidpointRounding.AwayFromZero).ToString() + " MPa";
                            Detensioning_In_Label.Text = Math.Round((TensioningPressureB * (1 + (DetensionProportion / 100M)) * 145.037738007M), 0, MidpointRounding.AwayFromZero).ToString() + " psi";
                            Detensioning_SI_Label.Text = Math.Round((TensioningPressureB * (1 + (DetensionProportion / 100M))), 0, MidpointRounding.AwayFromZero).ToString() + " MPa";
                            PressureA_In_Label.Text = Math.Round((TensioningPressureA * 145.037738007M), 0, MidpointRounding.AwayFromZero).ToString() + " psi";
                            PressureA_SI_Label.Text = Math.Round((TensioningPressureA), 2, MidpointRounding.AwayFromZero).ToString() + " MPa";
                            Pass1AppliedPressure_SequenceTabValue_In_Label.Text = Math.Round((TensioningPressureA * 145.037738007M), 0, MidpointRounding.AwayFromZero).ToString();
                            Pass1AppliedPressure_SequenceTabValue_SI_Label.Text = Math.Round((TensioningPressureA), 2, MidpointRounding.AwayFromZero).ToString();
                            Pass2Bolt_SequenceTabValue_Label.Text = "2";
                            Pass2AppliedPressure_SequenceTabValue_In_Label.Text = Math.Round((TensioningPressureB * 145.037738007M), 0, MidpointRounding.AwayFromZero).ToString();
                            Pass2AppliedPressure_SequenceTabValue_SI_Label.Text = Math.Round((TensioningPressureB), 2, MidpointRounding.AwayFromZero).ToString();
                            Pass3Bolt_SequenceTabValue_Label.Text = string.Empty;
                            Pass3AppliedPressure_SequenceTabValue_In_Label.Text = string.Empty;
                            Pass3AppliedPressure_SequenceTabValue_SI_Label.Text = string.Empty;
                            Pass4Bolt_SequenceTabValue_Label.Text = string.Empty;
                            Pass4AppliedPressure_SequenceTabValue_In_Label.Text = string.Empty;
                            Pass4AppliedPressure_SequenceTabValue_SI_Label.Text = string.Empty;
                            CheckPass1Bolt_SequenceTabValue_Label.Text = "1";
                            CheckPass1AppliedPressure_SequenceTabValue_In_Label.Text = Math.Round((TensioningPressureB * 145.037738007M), 0, MidpointRounding.AwayFromZero).ToString();
                            CheckPass1AppliedPressure_SequenceTabValue_SI_Label.Text = Math.Round((TensioningPressureB), 2, MidpointRounding.AwayFromZero).ToString();
                            CheckPass2Bolt_SequenceTabValue_Label.Text = "2";
                            CheckPass2AppliedPressure_SequenceTabValue_In_Label.Text = Math.Round((TensioningPressureB * 145.037738007M), 0, MidpointRounding.AwayFromZero).ToString();
                            CheckPass2AppliedPressure_SequenceTabValue_SI_Label.Text = Math.Round((TensioningPressureB), 2, MidpointRounding.AwayFromZero).ToString();
                            // For Torque only
                            Torque_SequenceTitle_Label.Visible = false;
                            Torque_SI_Sequence_Label.Visible = false;
                            Torque_In_Sequence_Label.Visible = false;
                            // Warnings that permitted values are exceeded
                            if ((ResidualBoltStress * LTFValue * (1 + (CrossLoad / 100M))) > (BoltYield * 0.95M))     // Only 95% of bolt yield value, pressure A & B allowed
                            {
                                TitleWarning_PictureBox.Visible = true;
                                PressureAWarning_PictureBox.Visible = true;
                            }
                            if ((ResidualBoltStress * LTFValue) > (BoltYield * 0.95M))     // Only 95% of bolt yield value, pressure A & B allowed
                            {
                                TitleWarning_PictureBox.Visible = true;
                                PressureBWarning_PictureBox.Visible = true;
                            }
                            if ((ResidualBoltStress) > (BoltYield * 0.95M))
                            {
                                TitleWarning_PictureBox.Visible = true;
                                ResidualStreeWarning_PictureBox.Visible = true;
                            }
                            if (TensioningPressureA > 150)                           // Only 150 MPa tensioning is allowed.
                            {
                                TitleWarning_PictureBox.Visible = true;
                                PressureA_Info_Warning_PictureBox.Visible = true;
                            }
                            if (TensioningPressureB > 150)                           // Only 150 MPa tensioning is allowed.
                            {
                                TitleWarning_PictureBox.Visible = true;
                                PressureB_Info_Warning_PictureBox.Visible = true;
                            }
                            if ((TensioningPressureB * (1 + (DetensionProportion / 100M))) > 150)
                            {
                                TitleWarning_PictureBox.Visible = true;
                                Deten_Info_Warning_PictureBox.Visible = true;
                            }
                        }
                        else
                        {
                            // Bolt Stress Tab Calculations for Bolt
                            T1BBoltStress_In_Label.Text = "";
                            T1BBoltStress_SI_Label.Text = "";
                            T1BBoltLoad_In_Label.Text = "";
                            T1BBoltLoad_SI_Label.Text = "";
                            T1BBoltYield_Per_Label.Text = "";

                            T1ABoltStress_In_Label.Text = "";
                            T1ABoltStress_SI_Label.Text = "";
                            T1ABoltLoad_In_Label.Text = "";
                            T1ABoltLoad_SI_Label.Text = "";
                            T1ABoltYield_SI_Label.Text = "";

                            T2RBoltStress_In_Label.Text = "";
                            T2RBoltStress_SI_Label.Text = "";
                            T2RBoltLoad_In_Label.Text = "";
                            T2RBoltLoad_SI_Label.Text = "";
                            T2RBoltYield_Per_Label.Text = "";

                            T3RBoltStress_In_Label.Text = Math.Round((ResidualBoltStress * 145.037738007M), 0, MidpointRounding.AwayFromZero).ToString();
                            T3RBoltStress_SI_Label.Text = ResidualBoltStress_TextBox.Text;
                            T3RBoltLoad_In_Label.Text = Math.Round((ResidualBoltStress * 145.037738007M * TensileStressArea_in / 2240M), 2, MidpointRounding.AwayFromZero).ToString();
                            T3RBoltLoad_SI_Label.Text = Math.Round((ResidualBoltStress * TensileStressArea_mm / 1000M), 2, MidpointRounding.AwayFromZero).ToString();
                            if (BoltYield > 0)
                            {
                                T3RBoltYield_Per_Label.Text = Math.Round(((ResidualBoltStress) * 100M / BoltYield), 2, MidpointRounding.AwayFromZero).ToString();
                            }
                            else
                            {
                                T3RBoltYield_Per_Label.Text = "";
                            }
                            DetenBoltStress_In_Label.Text = "";
                            DetenBoltStress_SI_Label.Text = "";
                            DetenBoltLoad_In_Label.Text = "";
                            DetenBoltLoad_SI_Label.Text = "";
                            DetenBoltYield_Per_Label.Text = "";

                            // Sequence Tab Calculations for Tool
                            PressureB_In_Label.Text = "";
                            PressureB_SI_Label.Text = "";
                            Detensioning_In_Label.Text = PressureB_In_Label.Text;
                            Detensioning_SI_Label.Text = PressureB_SI_Label.Text;
                            PressureA_In_Label.Text = string.Empty;
                            PressureA_SI_Label.Text = string.Empty;
                            Pass1AppliedPressure_SequenceTabValue_In_Label.Text = PressureA_In_Label.Text;
                            Pass1AppliedPressure_SequenceTabValue_SI_Label.Text = PressureA_SI_Label.Text;
                            Pass2Bolt_SequenceTabValue_Label.Text = "";
                            Pass2AppliedPressure_SequenceTabValue_In_Label.Text = PressureB_In_Label.Text;
                            Pass2AppliedPressure_SequenceTabValue_SI_Label.Text = PressureB_SI_Label.Text;
                            Pass3Bolt_SequenceTabValue_Label.Text = string.Empty;
                            Pass3AppliedPressure_SequenceTabValue_In_Label.Text = string.Empty;
                            Pass3AppliedPressure_SequenceTabValue_SI_Label.Text = string.Empty;
                            Pass4Bolt_SequenceTabValue_Label.Text = string.Empty;
                            Pass4AppliedPressure_SequenceTabValue_In_Label.Text = string.Empty;
                            Pass4AppliedPressure_SequenceTabValue_SI_Label.Text = string.Empty;
                            CheckPass1Bolt_SequenceTabValue_Label.Text = "";
                            CheckPass1AppliedPressure_SequenceTabValue_In_Label.Text = PressureB_In_Label.Text;
                            CheckPass1AppliedPressure_SequenceTabValue_SI_Label.Text = PressureB_SI_Label.Text;
                            CheckPass2Bolt_SequenceTabValue_Label.Text = "";
                            CheckPass2AppliedPressure_SequenceTabValue_In_Label.Text = PressureB_In_Label.Text;
                            CheckPass2AppliedPressure_SequenceTabValue_SI_Label.Text = PressureB_SI_Label.Text;
                            // For Torque only
                            Torque_SequenceTitle_Label.Visible = false;
                            Torque_SI_Sequence_Label.Visible = false;
                            Torque_In_Sequence_Label.Visible = false;
                        }
                    }
                    if (NumberOfBolts_TextBox.Text.Trim() != "")
                    {
                        int Bolts_Case01 = Convert.ToInt32(NumberOfBolts_TextBox.Text);
                        switch (Bolts_Case01)
                        {
                            case 4:
                                Sequence_PictureBox.Image = WizBolt.Properties.Resources.Flange050_04;
                                break;
                            case 8:
                                Sequence_PictureBox.Image = WizBolt.Properties.Resources.Flange050_08;
                                break;
                            case 12:
                                Sequence_PictureBox.Image = WizBolt.Properties.Resources.Flange050_12;
                                break;
                            case 16:
                                Sequence_PictureBox.Image = WizBolt.Properties.Resources.Flange050_16;
                                break;
                            case 20:
                                Sequence_PictureBox.Image = WizBolt.Properties.Resources.Flange050_20;
                                break;
                            case 24:
                                Sequence_PictureBox.Image = WizBolt.Properties.Resources.Flange050_24;
                                break;
                            case 28:
                                Sequence_PictureBox.Image = WizBolt.Properties.Resources.Flange050_28;
                                break;
                            case 32:
                                Sequence_PictureBox.Image = WizBolt.Properties.Resources.Flange050_32;
                                break;
                            case 36:
                                Sequence_PictureBox.Image = WizBolt.Properties.Resources.Flange050_36;
                                break;
                            case 40:
                                Sequence_PictureBox.Image = WizBolt.Properties.Resources.Flange050_40;
                                break;
                            case 44:
                                Sequence_PictureBox.Image = WizBolt.Properties.Resources.Flange050_44;
                                break;
                            case 48:
                                Sequence_PictureBox.Image = WizBolt.Properties.Resources.Flange050_48;
                                break;
                            case 52:
                                Sequence_PictureBox.Image = WizBolt.Properties.Resources.Flange050_52;
                                break;
                            case 56:
                                Sequence_PictureBox.Image = WizBolt.Properties.Resources.Flange050_56;
                                break;
                            case 60:
                                Sequence_PictureBox.Image = WizBolt.Properties.Resources.Flange050_60;
                                break;
                            default:
                                Sequence_PictureBox.Image = WizBolt.Properties.Resources.Sequence_half;
                                break;
                        }
                    }
                    else
                    {
                        Sequence_PictureBox.Image = null;
                    }
                    break;
                case 2:
                    Sequence_PictureBox.Image = WizBolt.Properties.Resources.Sequence_quarter;
                    if (UnitSystem_Button.Text == "SAE")
                    {
                        if (ToolHydraulicArea > 0)
                        {
                            // Bolt Stress Tab Calculations for Bolt
                            T1BBoltStress_In_Label.Text = Math.Round((ResidualBoltStress * LTFValue), 0, MidpointRounding.AwayFromZero).ToString();
                            T1BBoltStress_SI_Label.Text = Math.Round((ResidualBoltStress * LTFValue / 145.037738007M), 2, MidpointRounding.AwayFromZero).ToString();
                            T1BBoltLoad_In_Label.Text = Math.Round((ResidualBoltStress * LTFValue * TensileStressArea_in / 2240M), 2, MidpointRounding.AwayFromZero).ToString();
                            T1BBoltLoad_SI_Label.Text = Math.Round((ResidualBoltStress * TensileStressArea_mm * LTFValue / 145037.738007M), 2, MidpointRounding.AwayFromZero).ToString();
                            if (BoltYield > 0)
                            {
                                T1BBoltYield_Per_Label.Text = Math.Round(((ResidualBoltStress * LTFValue / 145.037738007M) * 100M / BoltYield), 2, MidpointRounding.AwayFromZero).ToString();
                            }
                            else
                            {
                                T1BBoltYield_Per_Label.Text = "";
                            }
                            T1ABoltStress_In_Label.Text = Math.Round((ResidualBoltStress * LTFValue * (1 + (CrossLoad / 100M))), 0, MidpointRounding.AwayFromZero).ToString(); ;
                            T1ABoltStress_SI_Label.Text = Math.Round((ResidualBoltStress * LTFValue * (1 + (CrossLoad / 100M)) / 145.037738007M), 2, MidpointRounding.AwayFromZero).ToString(); ;
                            T1ABoltLoad_In_Label.Text = Math.Round((ResidualBoltStress * LTFValue * (1 + (CrossLoad / 100M)) * TensileStressArea_in / 2240M), 2, MidpointRounding.AwayFromZero).ToString();
                            T1ABoltLoad_SI_Label.Text = Math.Round((ResidualBoltStress * (1 + (CrossLoad / 100M)) * TensileStressArea_mm * LTFValue / 145037.738007M), 2, MidpointRounding.AwayFromZero).ToString();
                            if (BoltYield > 0)
                            {
                                T1ABoltYield_SI_Label.Text = Math.Round(((ResidualBoltStress * (1 + (CrossLoad / 100M)) * LTFValue / 145.037738007M) * 100M / BoltYield), 2, MidpointRounding.AwayFromZero).ToString();
                            }
                            else
                            {
                                T1ABoltYield_SI_Label.Text = "";
                            }

                            T2RBoltStress_In_Label.Text = ResidualBoltStress_TextBox.Text;
                            T2RBoltStress_SI_Label.Text = Math.Round((ResidualBoltStress / 145.037738007M), 2, MidpointRounding.AwayFromZero).ToString();
                            T2RBoltLoad_In_Label.Text = Math.Round((ResidualBoltStress * TensileStressArea_in / 2240M), 2, MidpointRounding.AwayFromZero).ToString();
                            T2RBoltLoad_SI_Label.Text = Math.Round((ResidualBoltStress * TensileStressArea_mm / 145037.738007M), 2, MidpointRounding.AwayFromZero).ToString();
                            if (BoltYield > 0)
                            {
                                T2RBoltYield_Per_Label.Text = Math.Round(((ResidualBoltStress / 145.037738007M) * 100M / BoltYield), 2, MidpointRounding.AwayFromZero).ToString();
                            }
                            else
                            {
                                T2RBoltYield_Per_Label.Text = "";
                            }

                            T3RBoltStress_In_Label.Text = ResidualBoltStress_TextBox.Text;
                            T3RBoltStress_SI_Label.Text = Math.Round((ResidualBoltStress / 145.037738007M), 2, MidpointRounding.AwayFromZero).ToString();
                            T3RBoltLoad_In_Label.Text = Math.Round((ResidualBoltStress * TensileStressArea_in / 2240M), 2, MidpointRounding.AwayFromZero).ToString();
                            T3RBoltLoad_SI_Label.Text = Math.Round((ResidualBoltStress * TensileStressArea_mm / 145037.738007M), 2, MidpointRounding.AwayFromZero).ToString();
                            if (BoltYield > 0)
                            {
                                T3RBoltYield_Per_Label.Text = Math.Round(((ResidualBoltStress / 145.037738007M) * 100M / BoltYield), 2, MidpointRounding.AwayFromZero).ToString();
                            }
                            else
                            {
                                T3RBoltYield_Per_Label.Text = "";
                            }
                            DetenBoltStress_In_Label.Text = Math.Round((ResidualBoltStress * LTFValue * (1 + (DetensionProportion / 100M))), 0, MidpointRounding.AwayFromZero).ToString();
                            DetenBoltStress_SI_Label.Text = Math.Round((ResidualBoltStress * LTFValue * (1 + (DetensionProportion / 100M)) / 145.037738007M), 2, MidpointRounding.AwayFromZero).ToString();
                            DetenBoltLoad_In_Label.Text = Math.Round((ResidualBoltStress * LTFValue * (1 + (DetensionProportion / 100M)) * TensileStressArea_in / 2240M), 2, MidpointRounding.AwayFromZero).ToString();
                            DetenBoltLoad_SI_Label.Text = Math.Round((ResidualBoltStress * LTFValue * (1 + (DetensionProportion / 100M)) * TensileStressArea_mm / 145037.738007M), 2, MidpointRounding.AwayFromZero).ToString();
                            if (BoltYield > 0)
                            {
                                DetenBoltYield_Per_Label.Text = Math.Round(((ResidualBoltStress * LTFValue * (1 + (DetensionProportion / 100M)) / 145.037738007M) * 100M / BoltYield), 2, MidpointRounding.AwayFromZero).ToString();
                            }
                            else
                            {
                                DetenBoltYield_Per_Label.Text = "";
                            }

                            // Sequence Tab Calculations for Tool
                            TensioningPressureB = (ResidualBoltStress * TensileStressArea_in * LTFValue / (ToolHydraulicArea / 645.16M));
                            TensioningPressureA = TensioningPressureB * (1 + (CrossLoad / 100));
                            PressureB_In_Label.Text = Math.Round(TensioningPressureB, 0, MidpointRounding.AwayFromZero).ToString() + " psi";
                            PressureB_SI_Label.Text = Math.Round((TensioningPressureB / 145.037738007M), 2, MidpointRounding.AwayFromZero).ToString() + " MPa";
                            Detensioning_In_Label.Text = Math.Round((TensioningPressureB * (1 + (DetensionProportion / 100M))), 0, MidpointRounding.AwayFromZero).ToString() + " psi";
                            Detensioning_SI_Label.Text = Math.Round((TensioningPressureB * (1 + (DetensionProportion / 100M)) / 145.037738007M), 0, MidpointRounding.AwayFromZero).ToString() + " MPa";
                            PressureA_In_Label.Text = Math.Round(TensioningPressureA, 0, MidpointRounding.AwayFromZero).ToString() + " psi";
                            PressureA_SI_Label.Text = Math.Round((TensioningPressureA / 145.037738007M), 2, MidpointRounding.AwayFromZero).ToString() + " MPa";
                            Pass1AppliedPressure_SequenceTabValue_In_Label.Text = Math.Round(TensioningPressureA, 0, MidpointRounding.AwayFromZero).ToString();
                            Pass1AppliedPressure_SequenceTabValue_SI_Label.Text = Math.Round((TensioningPressureA / 145.037738007M), 2, MidpointRounding.AwayFromZero).ToString();
                            Pass2Bolt_SequenceTabValue_Label.Text = "3";
                            Pass2AppliedPressure_SequenceTabValue_In_Label.Text = Math.Round(TensioningPressureA, 0, MidpointRounding.AwayFromZero).ToString();
                            Pass2AppliedPressure_SequenceTabValue_SI_Label.Text = Math.Round((TensioningPressureA / 145.037738007M), 2, MidpointRounding.AwayFromZero).ToString();
                            Pass3Bolt_SequenceTabValue_Label.Text = "2";
                            Pass3AppliedPressure_SequenceTabValue_In_Label.Text = Math.Round(TensioningPressureB, 0, MidpointRounding.AwayFromZero).ToString();
                            Pass3AppliedPressure_SequenceTabValue_SI_Label.Text = Math.Round((TensioningPressureB / 145.037738007M), 2, MidpointRounding.AwayFromZero).ToString();
                            Pass4Bolt_SequenceTabValue_Label.Text = "4";
                            Pass4AppliedPressure_SequenceTabValue_In_Label.Text = Math.Round(TensioningPressureB, 0, MidpointRounding.AwayFromZero).ToString();
                            Pass4AppliedPressure_SequenceTabValue_SI_Label.Text = Math.Round((TensioningPressureB / 145.037738007M), 2, MidpointRounding.AwayFromZero).ToString();
                            CheckPass1Bolt_SequenceTabValue_Label.Text = "1";
                            CheckPass1AppliedPressure_SequenceTabValue_In_Label.Text = Math.Round(TensioningPressureB, 0, MidpointRounding.AwayFromZero).ToString();
                            CheckPass1AppliedPressure_SequenceTabValue_SI_Label.Text = Math.Round((TensioningPressureB / 145.037738007M), 2, MidpointRounding.AwayFromZero).ToString();
                            CheckPass2Bolt_SequenceTabValue_Label.Text = "2";
                            CheckPass2AppliedPressure_SequenceTabValue_In_Label.Text = Math.Round(TensioningPressureB, 0, MidpointRounding.AwayFromZero).ToString();
                            CheckPass2AppliedPressure_SequenceTabValue_SI_Label.Text = Math.Round((TensioningPressureB / 145.037738007M), 2, MidpointRounding.AwayFromZero).ToString();
                            // For Torque only
                            Torque_SequenceTitle_Label.Visible = false;
                            Torque_SI_Sequence_Label.Visible = false;
                            Torque_In_Sequence_Label.Visible = false;
                            // Warnings that permitted values are exceeded
                            if ((ResidualBoltStress * LTFValue * (1 + (CrossLoad / 100M)) / 145.037738007M) > (BoltYield * 0.95M))     // Only 95% of bolt yield value, pressure A & B allowed
                            {
                                TitleWarning_PictureBox.Visible = true;
                                PressureAWarning_PictureBox.Visible = true;
                            }
                            if ((ResidualBoltStress * LTFValue / 145.037738007M) > (BoltYield * 0.95M))     // Only 95% of bolt yield value, pressure A & B allowed
                            {
                                TitleWarning_PictureBox.Visible = true;
                                PressureBWarning_PictureBox.Visible = true;
                            }
                            if ((ResidualBoltStress / 145.037738007M) > (BoltYield * 0.95M))
                            {
                                TitleWarning_PictureBox.Visible = true;
                                ResidualStreeWarning_PictureBox.Visible = true;
                            }
                            if (TensioningPressureA > 21750M)                           // Only 21750 psi tensioning is allowed.
                            {
                                TitleWarning_PictureBox.Visible = true;
                                PressureA_Info_Warning_PictureBox.Visible = true;
                            }
                            if (TensioningPressureB > 21750M)                           // Only 21750 psi tensioning is allowed.
                            {
                                TitleWarning_PictureBox.Visible = true;
                                PressureB_Info_Warning_PictureBox.Visible = true;
                            }
                            if ((TensioningPressureB * (1 + (DetensionProportion / 100M))) > 21750M)
                            {
                                TitleWarning_PictureBox.Visible = true;
                                Deten_Info_Warning_PictureBox.Visible = true;
                            }

                        }
                        else
                        {
                            // Bolt Stress Tab Calculations for Bolt
                            T1BBoltStress_In_Label.Text = "";
                            T1BBoltStress_SI_Label.Text = "";
                            T1BBoltLoad_In_Label.Text = "";
                            T1BBoltLoad_SI_Label.Text = "";
                            T1BBoltYield_Per_Label.Text = "";

                            T1ABoltStress_In_Label.Text = "";
                            T1ABoltStress_SI_Label.Text = "";
                            T1ABoltLoad_In_Label.Text = "";
                            T1ABoltLoad_SI_Label.Text = "";
                            T1ABoltYield_SI_Label.Text = "";

                            T2RBoltStress_In_Label.Text = "";
                            T2RBoltStress_SI_Label.Text = "";
                            T2RBoltLoad_In_Label.Text = "";
                            T2RBoltLoad_SI_Label.Text = "";
                            T2RBoltYield_Per_Label.Text = "";

                            T3RBoltStress_In_Label.Text = ResidualBoltStress_TextBox.Text;
                            T3RBoltStress_SI_Label.Text = Math.Round((ResidualBoltStress / 145.037738007M), 2, MidpointRounding.AwayFromZero).ToString();
                            T3RBoltLoad_In_Label.Text = Math.Round((ResidualBoltStress * TensileStressArea_in / 2240M), 2, MidpointRounding.AwayFromZero).ToString();
                            T3RBoltLoad_SI_Label.Text = Math.Round((ResidualBoltStress * TensileStressArea_mm / 145037.738007M), 2, MidpointRounding.AwayFromZero).ToString();
                            if (BoltYield > 0)
                            {
                                T3RBoltYield_Per_Label.Text = Math.Round(((ResidualBoltStress / 145.037738007M) * 100M / BoltYield), 2, MidpointRounding.AwayFromZero).ToString();
                            }
                            else
                            {
                                T3RBoltYield_Per_Label.Text = "";
                            }
                            DetenBoltStress_In_Label.Text = "";
                            DetenBoltStress_SI_Label.Text = "";
                            DetenBoltLoad_In_Label.Text = "";
                            DetenBoltLoad_SI_Label.Text = "";
                            DetenBoltYield_Per_Label.Text = "";

                            // Sequence Tab Calculations for Tool
                            PressureB_In_Label.Text = "";
                            PressureB_SI_Label.Text = "";
                            Detensioning_In_Label.Text = PressureB_In_Label.Text;
                            Detensioning_SI_Label.Text = PressureB_SI_Label.Text;
                            PressureA_In_Label.Text = string.Empty;
                            PressureA_SI_Label.Text = string.Empty;
                            Pass1AppliedPressure_SequenceTabValue_In_Label.Text = PressureA_In_Label.Text;
                            Pass1AppliedPressure_SequenceTabValue_SI_Label.Text = PressureA_SI_Label.Text;
                            Pass2Bolt_SequenceTabValue_Label.Text = "";
                            Pass2AppliedPressure_SequenceTabValue_In_Label.Text = PressureA_In_Label.Text;
                            Pass2AppliedPressure_SequenceTabValue_SI_Label.Text = PressureA_SI_Label.Text;
                            Pass3AppliedPressure_SequenceTabValue_In_Label.Text = PressureB_In_Label.Text;
                            Pass3AppliedPressure_SequenceTabValue_SI_Label.Text = PressureB_SI_Label.Text;
                            Pass4AppliedPressure_SequenceTabValue_In_Label.Text = PressureB_In_Label.Text;
                            Pass4AppliedPressure_SequenceTabValue_SI_Label.Text = PressureB_SI_Label.Text;
                            CheckPass1Bolt_SequenceTabValue_Label.Text = "";
                            CheckPass1AppliedPressure_SequenceTabValue_In_Label.Text = PressureB_In_Label.Text;
                            CheckPass1AppliedPressure_SequenceTabValue_SI_Label.Text = PressureB_SI_Label.Text;
                            CheckPass2Bolt_SequenceTabValue_Label.Text = "";
                            CheckPass2AppliedPressure_SequenceTabValue_In_Label.Text = PressureB_In_Label.Text;
                            CheckPass2AppliedPressure_SequenceTabValue_SI_Label.Text = PressureB_SI_Label.Text;
                            // For Torque only
                            Torque_SequenceTitle_Label.Visible = false;
                            Torque_SI_Sequence_Label.Visible = false;
                            Torque_In_Sequence_Label.Visible = false;
                        }
                    }
                    else
                    {
                        if (ToolHydraulicArea > 0)
                        {
                            // Bolt Stress Tab Calculations for Bolt
                            T1BBoltStress_In_Label.Text = Math.Round((ResidualBoltStress * 145.037738007M * LTFValue), 0, MidpointRounding.AwayFromZero).ToString();
                            T1BBoltStress_SI_Label.Text = Math.Round((ResidualBoltStress * LTFValue), 2, MidpointRounding.AwayFromZero).ToString();
                            T1BBoltLoad_In_Label.Text = Math.Round((ResidualBoltStress * 145.037738007M * LTFValue * TensileStressArea_in / 2240M), 2, MidpointRounding.AwayFromZero).ToString();
                            T1BBoltLoad_SI_Label.Text = Math.Round((ResidualBoltStress * TensileStressArea_mm * LTFValue / 1000M), 2, MidpointRounding.AwayFromZero).ToString();
                            if (BoltYield > 0)
                            {
                                T1BBoltYield_Per_Label.Text = Math.Round(((ResidualBoltStress * LTFValue) * 100M / BoltYield), 2, MidpointRounding.AwayFromZero).ToString();
                            }
                            else
                            {
                                T1BBoltYield_Per_Label.Text = "";
                            }

                            T1ABoltStress_In_Label.Text = Math.Round((ResidualBoltStress * 145.037738007M * LTFValue * (1 + (CrossLoad / 100M))), 0, MidpointRounding.AwayFromZero).ToString();
                            T1ABoltStress_SI_Label.Text = Math.Round((ResidualBoltStress * LTFValue * (1 + (CrossLoad / 100M))), 2, MidpointRounding.AwayFromZero).ToString();
                            T1ABoltLoad_In_Label.Text = Math.Round((ResidualBoltStress * 145.037738007M * LTFValue * (1 + (CrossLoad / 100M)) * TensileStressArea_in / 2240M), 2, MidpointRounding.AwayFromZero).ToString();
                            T1ABoltLoad_SI_Label.Text = Math.Round((ResidualBoltStress * TensileStressArea_mm * LTFValue * (1 + (CrossLoad / 100M)) / 1000M), 2, MidpointRounding.AwayFromZero).ToString();
                            if (BoltYield > 0)
                            {
                                T1ABoltYield_SI_Label.Text = Math.Round(((ResidualBoltStress * LTFValue * (1 + (CrossLoad / 100M))) * 100M / BoltYield), 2, MidpointRounding.AwayFromZero).ToString();
                            }
                            else
                            {
                                T1ABoltYield_SI_Label.Text = "";
                            }

                            T2RBoltStress_In_Label.Text = Math.Round((ResidualBoltStress * 145.037738007M), 0, MidpointRounding.AwayFromZero).ToString();
                            T2RBoltStress_SI_Label.Text = ResidualBoltStress_TextBox.Text;
                            T2RBoltLoad_In_Label.Text = Math.Round((ResidualBoltStress * 145.037738007M * TensileStressArea_in / 2240M), 2, MidpointRounding.AwayFromZero).ToString();
                            T2RBoltLoad_SI_Label.Text = Math.Round((ResidualBoltStress * TensileStressArea_mm / 1000M), 2, MidpointRounding.AwayFromZero).ToString();
                            if (BoltYield > 0)
                            {
                                T2RBoltYield_Per_Label.Text = Math.Round(((ResidualBoltStress) * 100M / BoltYield), 2, MidpointRounding.AwayFromZero).ToString();
                            }
                            else
                            {
                                T2RBoltYield_Per_Label.Text = "";
                            }
                            T3RBoltStress_In_Label.Text = Math.Round((ResidualBoltStress * 145.037738007M), 0, MidpointRounding.AwayFromZero).ToString();
                            T3RBoltStress_SI_Label.Text = ResidualBoltStress_TextBox.Text;
                            T3RBoltLoad_In_Label.Text = Math.Round((ResidualBoltStress * 145.037738007M * TensileStressArea_in / 2240M), 2, MidpointRounding.AwayFromZero).ToString();
                            T3RBoltLoad_SI_Label.Text = Math.Round((ResidualBoltStress * TensileStressArea_mm / 1000M), 2, MidpointRounding.AwayFromZero).ToString();
                            if (BoltYield > 0)
                            {
                                T3RBoltYield_Per_Label.Text = Math.Round(((ResidualBoltStress) * 100M / BoltYield), 2, MidpointRounding.AwayFromZero).ToString();
                            }
                            else
                            {
                                T3RBoltYield_Per_Label.Text = "";
                            }
                            DetenBoltStress_In_Label.Text = Math.Round((ResidualBoltStress * 145.037738007M * LTFValue * (1 + (DetensionProportion / 100M))), 0, MidpointRounding.AwayFromZero).ToString();
                            DetenBoltStress_SI_Label.Text = Math.Round((ResidualBoltStress * LTFValue * (1 + (DetensionProportion / 100M))), 2, MidpointRounding.AwayFromZero).ToString();
                            DetenBoltLoad_In_Label.Text = Math.Round((ResidualBoltStress * 145.037738007M * LTFValue * (1 + (DetensionProportion / 100M)) * TensileStressArea_in / 2240M), 2, MidpointRounding.AwayFromZero).ToString();
                            DetenBoltLoad_SI_Label.Text = Math.Round((ResidualBoltStress * LTFValue * (1 + (DetensionProportion / 100M)) * TensileStressArea_mm / 1000M), 2, MidpointRounding.AwayFromZero).ToString();
                            if (BoltYield > 0)
                            {
                                DetenBoltYield_Per_Label.Text = Math.Round(((ResidualBoltStress * LTFValue * (1 + (DetensionProportion / 100M))) * 100M / BoltYield), 2, MidpointRounding.AwayFromZero).ToString();
                            }
                            else
                            {
                                DetenBoltYield_Per_Label.Text = "";
                            }

                            // Sequence Tab Calculations for Tool
                            if (ToolHydraulicArea > 0)
                            {
                                TensioningPressureB = (ResidualBoltStress * TensileStressArea_mm * LTFValue / ToolHydraulicArea);
                            }
                            TensioningPressureA = TensioningPressureB * (1 + (CrossLoad / 100));
                            PressureB_In_Label.Text = Math.Round((TensioningPressureB * 145.037738007M), 0, MidpointRounding.AwayFromZero).ToString() + " psi";
                            PressureB_SI_Label.Text = Math.Round((TensioningPressureB), 2, MidpointRounding.AwayFromZero).ToString() + " MPa";
                            Detensioning_In_Label.Text = Math.Round((TensioningPressureB * (1 + (DetensionProportion / 100M)) * 145.037738007M), 0, MidpointRounding.AwayFromZero).ToString() + " psi";
                            Detensioning_SI_Label.Text = Math.Round((TensioningPressureB * (1 + (DetensionProportion / 100M))), 0, MidpointRounding.AwayFromZero).ToString() + " MPa";
                            PressureA_In_Label.Text = Math.Round((TensioningPressureA * 145.037738007M), 0, MidpointRounding.AwayFromZero).ToString() + " psi";
                            PressureA_SI_Label.Text = Math.Round((TensioningPressureA), 2, MidpointRounding.AwayFromZero).ToString() + " MPa";
                            Pass1AppliedPressure_SequenceTabValue_In_Label.Text = Math.Round((TensioningPressureA * 145.037738007M), 0, MidpointRounding.AwayFromZero).ToString();
                            Pass1AppliedPressure_SequenceTabValue_SI_Label.Text = Math.Round((TensioningPressureA), 2, MidpointRounding.AwayFromZero).ToString();
                            Pass2Bolt_SequenceTabValue_Label.Text = "3";
                            Pass2AppliedPressure_SequenceTabValue_In_Label.Text = Math.Round((TensioningPressureA * 145.037738007M), 0, MidpointRounding.AwayFromZero).ToString();
                            Pass2AppliedPressure_SequenceTabValue_SI_Label.Text = Math.Round((TensioningPressureA), 2, MidpointRounding.AwayFromZero).ToString();
                            Pass3Bolt_SequenceTabValue_Label.Text = "2";
                            Pass3AppliedPressure_SequenceTabValue_In_Label.Text = Math.Round((TensioningPressureB * 145.037738007M), 0, MidpointRounding.AwayFromZero).ToString();
                            Pass3AppliedPressure_SequenceTabValue_SI_Label.Text = Math.Round((TensioningPressureB), 2, MidpointRounding.AwayFromZero).ToString();
                            Pass4Bolt_SequenceTabValue_Label.Text = "4";
                            Pass4AppliedPressure_SequenceTabValue_In_Label.Text = Math.Round((TensioningPressureB * 145.037738007M), 0, MidpointRounding.AwayFromZero).ToString();
                            Pass4AppliedPressure_SequenceTabValue_SI_Label.Text = Math.Round((TensioningPressureB), 2, MidpointRounding.AwayFromZero).ToString();
                            CheckPass1Bolt_SequenceTabValue_Label.Text = "1";
                            CheckPass1AppliedPressure_SequenceTabValue_In_Label.Text = Math.Round((TensioningPressureB * 145.037738007M), 0, MidpointRounding.AwayFromZero).ToString();
                            CheckPass1AppliedPressure_SequenceTabValue_SI_Label.Text = Math.Round((TensioningPressureB), 2, MidpointRounding.AwayFromZero).ToString();
                            CheckPass2Bolt_SequenceTabValue_Label.Text = "2";
                            CheckPass2AppliedPressure_SequenceTabValue_In_Label.Text = Math.Round((TensioningPressureB * 145.037738007M), 0, MidpointRounding.AwayFromZero).ToString();
                            CheckPass2AppliedPressure_SequenceTabValue_SI_Label.Text = Math.Round((TensioningPressureB), 2, MidpointRounding.AwayFromZero).ToString();
                            // For Torque only
                            Torque_SequenceTitle_Label.Visible = false;
                            Torque_SI_Sequence_Label.Visible = false;
                            Torque_In_Sequence_Label.Visible = false;
                            // Warnings that permitted values are exceeded
                            if ((ResidualBoltStress * LTFValue * (1 + (CrossLoad / 100M))) > (BoltYield * 0.95M))     // Only 95% of bolt yield value, pressure A & B allowed
                            {
                                TitleWarning_PictureBox.Visible = true;
                                PressureAWarning_PictureBox.Visible = true;
                            }
                            if ((ResidualBoltStress * LTFValue) > (BoltYield * 0.95M))     // Only 95% of bolt yield value, pressure A & B allowed
                            {
                                TitleWarning_PictureBox.Visible = true;
                                PressureBWarning_PictureBox.Visible = true;
                            }
                            if ((ResidualBoltStress) > (BoltYield * 0.95M))
                            {
                                TitleWarning_PictureBox.Visible = true;
                                ResidualStreeWarning_PictureBox.Visible = true;
                            }
                            if (TensioningPressureA > 150)                           // Only 150 MPa tensioning is allowed.
                            {
                                TitleWarning_PictureBox.Visible = true;
                                PressureA_Info_Warning_PictureBox.Visible = true;
                            }
                            if (TensioningPressureB > 150)                           // Only 150 MPa tensioning is allowed.
                            {
                                TitleWarning_PictureBox.Visible = true;
                                PressureB_Info_Warning_PictureBox.Visible = true;
                            }
                            if ((TensioningPressureB * (1 + (DetensionProportion / 100M))) > 150)
                            {
                                TitleWarning_PictureBox.Visible = true;
                                Deten_Info_Warning_PictureBox.Visible = true;
                            }

                        }
                        else
                        {
                            T1BBoltStress_In_Label.Text = "";
                            T1BBoltStress_SI_Label.Text = "";
                            T1BBoltLoad_In_Label.Text = "";
                            T1BBoltLoad_SI_Label.Text = "";
                            T1BBoltYield_Per_Label.Text = "";

                            T1ABoltStress_In_Label.Text = "";
                            T1ABoltStress_SI_Label.Text = "";
                            T1ABoltLoad_In_Label.Text = "";
                            T1ABoltLoad_SI_Label.Text = "";
                            T1ABoltYield_SI_Label.Text = "";

                            T2RBoltStress_In_Label.Text = "";
                            T2RBoltStress_SI_Label.Text = "";
                            T2RBoltLoad_In_Label.Text = "";
                            T2RBoltLoad_SI_Label.Text = "";
                            T2RBoltYield_Per_Label.Text = "";

                            T3RBoltStress_In_Label.Text = Math.Round((ResidualBoltStress * 145.037738007M), 0, MidpointRounding.AwayFromZero).ToString();
                            T3RBoltStress_SI_Label.Text = ResidualBoltStress_TextBox.Text;
                            T3RBoltLoad_In_Label.Text = Math.Round((ResidualBoltStress * 145.037738007M * TensileStressArea_in / 2240M), 2, MidpointRounding.AwayFromZero).ToString();
                            T3RBoltLoad_SI_Label.Text = Math.Round((ResidualBoltStress * TensileStressArea_mm / 1000M), 2, MidpointRounding.AwayFromZero).ToString();
                            if (BoltYield > 0)
                            {
                                T3RBoltYield_Per_Label.Text = Math.Round(((ResidualBoltStress) * 100M / BoltYield), 2, MidpointRounding.AwayFromZero).ToString();
                            }
                            else
                            {
                                T3RBoltYield_Per_Label.Text = "";
                            }
                            DetenBoltStress_In_Label.Text = "";
                            DetenBoltStress_SI_Label.Text = "";
                            DetenBoltLoad_In_Label.Text = "";
                            DetenBoltLoad_SI_Label.Text = "";
                            DetenBoltYield_Per_Label.Text = "";

                            // Sequence Tab Calculations for Tool
                            PressureB_In_Label.Text = "";
                            PressureB_SI_Label.Text = "";
                            Detensioning_In_Label.Text = PressureB_In_Label.Text;
                            Detensioning_SI_Label.Text = PressureB_SI_Label.Text;
                            PressureA_In_Label.Text = string.Empty;
                            PressureA_SI_Label.Text = string.Empty;
                            Pass1AppliedPressure_SequenceTabValue_In_Label.Text = PressureA_In_Label.Text;
                            Pass1AppliedPressure_SequenceTabValue_SI_Label.Text = PressureA_SI_Label.Text;
                            Pass2Bolt_SequenceTabValue_Label.Text = "";
                            Pass2AppliedPressure_SequenceTabValue_In_Label.Text = PressureA_In_Label.Text;
                            Pass2AppliedPressure_SequenceTabValue_SI_Label.Text = PressureA_SI_Label.Text;
                            Pass3AppliedPressure_SequenceTabValue_In_Label.Text = PressureB_In_Label.Text;
                            Pass3AppliedPressure_SequenceTabValue_SI_Label.Text = PressureB_SI_Label.Text;
                            Pass4AppliedPressure_SequenceTabValue_In_Label.Text = PressureB_In_Label.Text;
                            Pass4AppliedPressure_SequenceTabValue_SI_Label.Text = PressureB_SI_Label.Text;
                            CheckPass1Bolt_SequenceTabValue_Label.Text = "";
                            CheckPass1AppliedPressure_SequenceTabValue_In_Label.Text = PressureB_In_Label.Text;
                            CheckPass1AppliedPressure_SequenceTabValue_SI_Label.Text = PressureB_SI_Label.Text;
                            CheckPass2Bolt_SequenceTabValue_Label.Text = "";
                            CheckPass2AppliedPressure_SequenceTabValue_In_Label.Text = PressureB_In_Label.Text;
                            CheckPass2AppliedPressure_SequenceTabValue_SI_Label.Text = PressureB_SI_Label.Text;
                            // For Torque only
                            Torque_SequenceTitle_Label.Visible = false;
                            Torque_SI_Sequence_Label.Visible = false;
                            Torque_In_Sequence_Label.Visible = false;
                        }
                    }
                    if (NumberOfBolts_TextBox.Text.Trim() != "")
                    {
                        int Bolts_Case02 = Convert.ToInt32(NumberOfBolts_TextBox.Text);
                        switch (Bolts_Case02)
                        {
                            case 4:
                                Sequence_PictureBox.Image = WizBolt.Properties.Resources.Flange025_04;
                                break;
                            case 8:
                                Sequence_PictureBox.Image = WizBolt.Properties.Resources.Flange025_08;
                                break;
                            case 12:
                                Sequence_PictureBox.Image = WizBolt.Properties.Resources.Flange025_12;
                                break;
                            case 16:
                                Sequence_PictureBox.Image = WizBolt.Properties.Resources.Flange025_16;
                                break;
                            case 20:
                                Sequence_PictureBox.Image = WizBolt.Properties.Resources.Flange025_20;
                                break;
                            case 24:
                                Sequence_PictureBox.Image = WizBolt.Properties.Resources.Flange025_24;
                                break;
                            case 28:
                                Sequence_PictureBox.Image = WizBolt.Properties.Resources.Flange025_28;
                                break;
                            case 32:
                                Sequence_PictureBox.Image = WizBolt.Properties.Resources.Flange025_32;
                                break;
                            case 36:
                                Sequence_PictureBox.Image = WizBolt.Properties.Resources.Flange025_36;
                                break;
                            case 40:
                                Sequence_PictureBox.Image = WizBolt.Properties.Resources.Flange025_40;
                                break;
                            case 44:
                                Sequence_PictureBox.Image = WizBolt.Properties.Resources.Flange025_44;
                                break;
                            case 48:
                                Sequence_PictureBox.Image = WizBolt.Properties.Resources.Flange025_48;
                                break;
                            case 52:
                                Sequence_PictureBox.Image = WizBolt.Properties.Resources.Flange025_52;
                                break;
                            case 56:
                                Sequence_PictureBox.Image = WizBolt.Properties.Resources.Flange025_56;
                                break;
                            case 60:
                                Sequence_PictureBox.Image = WizBolt.Properties.Resources.Flange025_60;
                                break;
                            default:
                                Sequence_PictureBox.Image = WizBolt.Properties.Resources.Sequence_quarter;
                                break;
                        }
                    }
                    else
                    {
                        Sequence_PictureBox.Image = null;
                    }
                    break;
                case 3:
                    if (UnitSystem_Button.Text == "SAE")
                    {
                        // Bolt Stress Tab Calculations for Bolt
                        T1BBoltStress_In_Label.Text = "";
                        T1BBoltStress_SI_Label.Text = "";
                        T1BBoltLoad_In_Label.Text = "";
                        T1BBoltLoad_SI_Label.Text = "";
                        T1BBoltYield_Per_Label.Text = "";

                        T1ABoltStress_In_Label.Text = "";
                        T1ABoltStress_SI_Label.Text = "";
                        T1ABoltLoad_In_Label.Text = "";
                        T1ABoltLoad_SI_Label.Text = "";
                        T1ABoltYield_SI_Label.Text = "";

                        T2RBoltStress_In_Label.Text = "";
                        T2RBoltStress_SI_Label.Text = "";
                        T2RBoltLoad_In_Label.Text = "";
                        T2RBoltLoad_SI_Label.Text = "";
                        T2RBoltYield_Per_Label.Text = "";

                        T3RBoltStress_In_Label.Text = ResidualBoltStress_TextBox.Text;
                        T3RBoltStress_SI_Label.Text = Math.Round((ResidualBoltStress / 145.037738007M), 2, MidpointRounding.AwayFromZero).ToString();
                        T3RBoltLoad_In_Label.Text = Math.Round((ResidualBoltStress * TensileStressArea_in / 2240M), 2, MidpointRounding.AwayFromZero).ToString();
                        T3RBoltLoad_SI_Label.Text = Math.Round((ResidualBoltStress * TensileStressArea_mm / 145037.738007M), 2, MidpointRounding.AwayFromZero).ToString();
                        if (BoltYield > 0)
                        {
                            T3RBoltYield_Per_Label.Text = Math.Round(((ResidualBoltStress / 145.037738007M) * 100M / BoltYield), 2, MidpointRounding.AwayFromZero).ToString();
                        }
                        else
                        {
                            T3RBoltYield_Per_Label.Text = "";
                        }
                        DetenBoltStress_In_Label.Text = "";
                        DetenBoltStress_SI_Label.Text = "";
                        DetenBoltLoad_In_Label.Text = "";
                        DetenBoltLoad_SI_Label.Text = "";
                        DetenBoltYield_Per_Label.Text = "";
                    }
                    else
                    {
                        // Bolt Stress Tab Calculations for Bolt
                        T1BBoltStress_In_Label.Text = "";
                        T1BBoltStress_SI_Label.Text = "";
                        T1BBoltLoad_In_Label.Text = "";
                        T1BBoltLoad_SI_Label.Text = "";
                        T1BBoltYield_Per_Label.Text = "";

                        T1ABoltStress_In_Label.Text = "";
                        T1ABoltStress_SI_Label.Text = "";
                        T1ABoltLoad_In_Label.Text = "";
                        T1ABoltLoad_SI_Label.Text = "";
                        T1ABoltYield_SI_Label.Text = "";

                        T2RBoltStress_In_Label.Text = "";
                        T2RBoltStress_SI_Label.Text = "";
                        T2RBoltLoad_In_Label.Text = "";
                        T2RBoltLoad_SI_Label.Text = "";
                        T2RBoltYield_Per_Label.Text = "";

                        T3RBoltStress_In_Label.Text = Math.Round((ResidualBoltStress * 145.037738007M), 0, MidpointRounding.AwayFromZero).ToString();
                        T3RBoltStress_SI_Label.Text = ResidualBoltStress_TextBox.Text;
                        T3RBoltLoad_In_Label.Text = Math.Round((ResidualBoltStress * 145.037738007M * TensileStressArea_in / 2240M), 2, MidpointRounding.AwayFromZero).ToString();
                        T3RBoltLoad_SI_Label.Text = Math.Round((ResidualBoltStress * TensileStressArea_mm / 1000M), 2, MidpointRounding.AwayFromZero).ToString();
                        if (BoltYield > 0)
                        {
                            T3RBoltYield_Per_Label.Text = Math.Round(((ResidualBoltStress) * 100M / BoltYield), 2, MidpointRounding.AwayFromZero).ToString();
                        }
                        else
                        {
                            T3RBoltYield_Per_Label.Text = "";
                        }
                        DetenBoltStress_In_Label.Text = "";
                        DetenBoltStress_SI_Label.Text = "";
                        DetenBoltLoad_In_Label.Text = "";
                        DetenBoltLoad_SI_Label.Text = "";
                        DetenBoltYield_Per_Label.Text = "";
                    }
                    // Sequence Tab Calculations for Tool
                    PressureB_In_Label.Text = "";
                    PressureB_SI_Label.Text = "";
                    Detensioning_In_Label.Text = PressureB_In_Label.Text;
                    Detensioning_SI_Label.Text = PressureB_SI_Label.Text;
                    PressureA_In_Label.Text = string.Empty;
                    PressureA_SI_Label.Text = string.Empty;
                    Pass1AppliedPressure_SequenceTabValue_In_Label.Text = PressureA_In_Label.Text;
                    Pass1AppliedPressure_SequenceTabValue_SI_Label.Text = PressureA_SI_Label.Text;
                    Pass2Bolt_SequenceTabValue_Label.Text = "";
                    Pass2AppliedPressure_SequenceTabValue_In_Label.Text = PressureA_In_Label.Text;
                    Pass2AppliedPressure_SequenceTabValue_SI_Label.Text = PressureA_SI_Label.Text;
                    Pass3AppliedPressure_SequenceTabValue_In_Label.Text = PressureB_In_Label.Text;
                    Pass3AppliedPressure_SequenceTabValue_SI_Label.Text = PressureB_SI_Label.Text;
                    Pass4AppliedPressure_SequenceTabValue_In_Label.Text = PressureB_In_Label.Text;
                    Pass4AppliedPressure_SequenceTabValue_SI_Label.Text = PressureB_SI_Label.Text;
                    CheckPass1Bolt_SequenceTabValue_Label.Text = "";
                    CheckPass1AppliedPressure_SequenceTabValue_In_Label.Text = PressureB_In_Label.Text;
                    CheckPass1AppliedPressure_SequenceTabValue_SI_Label.Text = PressureB_SI_Label.Text;
                    CheckPass2Bolt_SequenceTabValue_Label.Text = "";
                    CheckPass2AppliedPressure_SequenceTabValue_In_Label.Text = PressureB_In_Label.Text;
                    CheckPass2AppliedPressure_SequenceTabValue_SI_Label.Text = PressureB_SI_Label.Text;
                    // For Torque only
                    Torque_SequenceTitle_Label.Visible = true;
                    Torque_SI_Sequence_Label.Visible = true;
                    Torque_In_Sequence_Label.Visible = true;

                    if (NumberOfBolts_TextBox.Text.Trim() != "")
                    {
                        int Bolts_Case03 = Convert.ToInt32(NumberOfBolts_TextBox.Text);
                        switch (Bolts_Case03)
                        {
                            case 4:
                                Sequence_PictureBox.Image = WizBolt.Properties.Resources.FlangeTorque_04;
                                break;
                            case 8:
                                Sequence_PictureBox.Image = WizBolt.Properties.Resources.FlangeTorque_08;
                                break;
                            case 12:
                                Sequence_PictureBox.Image = WizBolt.Properties.Resources.FlangeTorque_12;
                                break;
                            case 16:
                                Sequence_PictureBox.Image = WizBolt.Properties.Resources.FlangeTorque_16;
                                break;
                            case 20:
                                Sequence_PictureBox.Image = WizBolt.Properties.Resources.FlangeTorque_20;
                                break;
                            case 24:
                                Sequence_PictureBox.Image = WizBolt.Properties.Resources.FlangeTorque_24;
                                break;
                            case 28:
                                Sequence_PictureBox.Image = WizBolt.Properties.Resources.FlangeTorque_28;
                                break;
                            case 32:
                                Sequence_PictureBox.Image = WizBolt.Properties.Resources.FlangeTorque_32;
                                break;
                            case 36:
                                Sequence_PictureBox.Image = WizBolt.Properties.Resources.FlangeTorque_36;
                                break;
                            case 40:
                                Sequence_PictureBox.Image = WizBolt.Properties.Resources.FlangeTorque_40;
                                break;
                            case 44:
                                Sequence_PictureBox.Image = WizBolt.Properties.Resources.FlangeTorque_44;
                                break;
                            case 48:
                                Sequence_PictureBox.Image = WizBolt.Properties.Resources.FlangeTorque_48;
                                break;
                            case 52:
                                Sequence_PictureBox.Image = WizBolt.Properties.Resources.FlangeTorque_52;
                                break;
                            case 56:
                                Sequence_PictureBox.Image = WizBolt.Properties.Resources.FlangeTorque_56;
                                break;
                            case 60:
                                Sequence_PictureBox.Image = WizBolt.Properties.Resources.FlangeTorque_60;
                                break;
                            default:
                                Sequence_PictureBox.Image = null;
                                break;
                        }
                    }
                    else
                    {
                        Sequence_PictureBox.Image = null;
                    }
                    break;
            }
            // Graph
            DrawGraph();
        }

        private void Detensioning_TextBox_Leave(object sender, EventArgs e)
        {
            decimal DetensionProportion = 0M;
            decimal BoltYield = 0M;
            decimal ResidualBoltStress = 0M;
            decimal LTFValue = 1.11M;
            decimal TensileStressArea_in = 0M;
            decimal TensileStressArea_mm = 0M;
            decimal TensioningPressureB = 0M;
            decimal ToolPressureArea = 0M;

            Deten_Info_Warning_PictureBox.Visible = false;          // If detension warning is on then first close it. The last part check whether warning need to display.

            if ((BoltYield_SI_Label.Text.Trim() != "") && (BoltYield_SI_Label.Text.Trim().Length > 3))
            {
                string YieldValue = BoltYield_SI_Label.Text.Substring(0, BoltYield_SI_Label.Text.Length - 3);
                BoltYield = Convert.ToDecimal(YieldValue);
            }
            if ((ResidualBoltStress_TextBox.Text.Trim() != "") && (ResidualBoltStress_TextBox.Text.Trim().Length > 0))
            {
                ResidualBoltStress = Convert.ToDecimal(ResidualBoltStress_TextBox.Text.Trim());
            }
             if ((TensileStressArea_In_Label.Text.Trim() != "") && (TensileStressArea_In_Label.Text.Trim().Length > 5))
            {
                string StressAreaValue_in = TensileStressArea_In_Label.Text.Substring(0, TensileStressArea_In_Label.Text.Length - 5);
                TensileStressArea_in = Convert.ToDecimal(StressAreaValue_in);
            }
            if ((TensileStressArea_SI_Label.Text.Trim() != "") && (TensileStressArea_SI_Label.Text.Trim().Length > 5))
            {
                string StressAreaValue_mm = TensileStressArea_SI_Label.Text.Substring(0, TensileStressArea_SI_Label.Text.Length - 5);
                TensileStressArea_mm = Convert.ToDecimal(StressAreaValue_mm);
            }
            
            if ((LTF_Label.Text.Trim() != "") && (LTF_Label.Text.Trim().Length > 0))
            {
                LTFValue = Convert.ToDecimal(LTF_Label.Text.Trim());
            }
            if ((ToolPressureArea_mm_Label.Text.Trim() != "") && (ToolPressureArea_mm_Label.Text.Trim().Length > 6))
            {
                string ToolArea = ToolPressureArea_mm_Label.Text.Substring(0, ToolPressureArea_mm_Label.Text.Length - 6);
                ToolPressureArea = Convert.ToDecimal(ToolArea);
            }
            
            if ((Detensioning_TextBox.Text.Trim() != "") && (Detensioning_TextBox.Text.Trim().Length > 0))
            {
                DetensionProportion = Convert.ToDecimal(Detensioning_TextBox.Text.Trim());
            }
            if (UnitSystem_Button.Text == "SAE")
            {
                ResidualBoltStress = ResidualBoltStress / 145.037738007M;
            }

            DetenBoltStress_In_Label.Text = Math.Round((ResidualBoltStress * LTFValue * (1 + (DetensionProportion / 100M)) *  145.037738007M), 0, MidpointRounding.AwayFromZero).ToString();
            DetenBoltStress_SI_Label.Text = Math.Round((ResidualBoltStress * LTFValue * (1 + (DetensionProportion / 100M))), 2, MidpointRounding.AwayFromZero).ToString();
            DetenBoltLoad_In_Label.Text = Math.Round((ResidualBoltStress * LTFValue * TensileStressArea_in * (1 + (DetensionProportion / 100M)) * 145037.738007M / 2240M), 2, MidpointRounding.AwayFromZero).ToString();
            DetenBoltLoad_SI_Label.Text = Math.Round((ResidualBoltStress * TensileStressArea_mm * LTFValue * (1 + (DetensionProportion / 100M))), 2, MidpointRounding.AwayFromZero).ToString();
            if (BoltYield > 0)
            {
                DetenBoltYield_Per_Label.Text = Math.Round(((ResidualBoltStress * LTFValue * (1 + (DetensionProportion / 100M))) * 100M / BoltYield), 2, MidpointRounding.AwayFromZero).ToString();
            }
            else
            {
                DetenBoltYield_Per_Label.Text = "";
            }
            TensioningPressureB = (ResidualBoltStress * TensileStressArea_mm * LTFValue / (ToolPressureArea));
            Detensioning_In_Label.Text = Math.Round((TensioningPressureB * (1 + (DetensionProportion / 100M)) * 145.037738007M), 0, MidpointRounding.AwayFromZero).ToString() + " psi";
            Detensioning_SI_Label.Text = Math.Round((TensioningPressureB * (1 + (DetensionProportion / 100M))), 2, MidpointRounding.AwayFromZero).ToString() + " MPa";

            if ((TensioningPressureB * (1 + (DetensionProportion / 100M))) > 150)
            {
                TitleWarning_PictureBox.Visible = true;
                Deten_Info_Warning_PictureBox.Visible = true;
            }

            // Graph
            DrawGraph();
        }

        private void CrossLoading_TextBox_Leave(object sender, EventArgs e)
        {
            decimal CrossLoad = 0M;
            decimal BoltYield = 0M;
            decimal ResidualBoltStress = 0M;
            decimal LTFValue = 1.11M;
            decimal TensileStressArea_in = 0M;
            decimal TensileStressArea_mm = 0M;
            decimal TensioningPressureB = 0M;
            decimal TensioningPressureA = 0M;
            decimal ToolPressureArea = 0M;

            // Switch off the warnings. These are displayed if it is required at respective places.
            
            PressureAWarning_PictureBox.Visible = false;
            PressureBWarning_PictureBox.Visible = false;
            ResidualStreeWarning_PictureBox.Visible = false;

            PressureA_Info_Warning_PictureBox.Visible = false;
            PressureB_Info_Warning_PictureBox.Visible = false;
            


            if ((BoltYield_SI_Label.Text.Trim() != "") && (BoltYield_SI_Label.Text.Trim().Length > 3))
            {
                string YieldValue = BoltYield_SI_Label.Text.Substring(0, BoltYield_SI_Label.Text.Length - 3);
                BoltYield = Convert.ToDecimal(YieldValue);
            }
            if ((ResidualBoltStress_TextBox.Text.Trim() != "") && (ResidualBoltStress_TextBox.Text.Trim().Length > 0))
            {
                ResidualBoltStress = Convert.ToDecimal(ResidualBoltStress_TextBox.Text.Trim());
            }
            if ((TensileStressArea_In_Label.Text.Trim() != "") && (TensileStressArea_In_Label.Text.Trim().Length > 5))
            {
                string StressAreaValue_in = TensileStressArea_In_Label.Text.Substring(0, TensileStressArea_In_Label.Text.Length - 5);
                TensileStressArea_in = Convert.ToDecimal(StressAreaValue_in);
            }
            if ((TensileStressArea_SI_Label.Text.Trim() != "") && (TensileStressArea_SI_Label.Text.Trim().Length > 5))
            {
                string StressAreaValue_mm = TensileStressArea_SI_Label.Text.Substring(0, TensileStressArea_SI_Label.Text.Length - 5);
                TensileStressArea_mm = Convert.ToDecimal(StressAreaValue_mm);
            }

            if ((LTF_Label.Text.Trim() != "") && (LTF_Label.Text.Trim().Length > 0))
            {
                LTFValue = Convert.ToDecimal(LTF_Label.Text.Trim());
            }
            if ((ToolPressureArea_mm_Label.Text.Trim() != "") && (ToolPressureArea_mm_Label.Text.Trim().Length > 6))
            {
                string ToolArea = ToolPressureArea_mm_Label.Text.Substring(0, ToolPressureArea_mm_Label.Text.Length - 6);
                ToolPressureArea = Convert.ToDecimal(ToolArea);
            }

            if ((CrossLoading_TextBox.Text.Trim() != "") && (CrossLoading_TextBox.Text.Trim().Length > 0))
            {
                CrossLoad = Convert.ToDecimal(CrossLoading_TextBox.Text.Trim());
            }
            if (UnitSystem_Button.Text == "SAE")
            {
                ResidualBoltStress = ResidualBoltStress / 145.037738007M;
            }

            int Bolt_Tool_Proportion = BoltTool_ComboBox.SelectedIndex;
            switch (Bolt_Tool_Proportion)
            {
                case 0:
                    break;
                case 1:
                    T1ABoltStress_In_Label.Text = Math.Round((ResidualBoltStress * LTFValue * (1 + (CrossLoad / 100M)) * 145.037738007M), 0, MidpointRounding.AwayFromZero).ToString(); ;
                    T1ABoltStress_SI_Label.Text = Math.Round((ResidualBoltStress * LTFValue * (1 + (CrossLoad / 100M))), 2, MidpointRounding.AwayFromZero).ToString(); ;
                    T1ABoltLoad_In_Label.Text = Math.Round((ResidualBoltStress * LTFValue * (1 + (CrossLoad / 100M)) * TensileStressArea_in * 145.037738007M / 2240M), 2, MidpointRounding.AwayFromZero).ToString();
                    T1ABoltLoad_SI_Label.Text = Math.Round((ResidualBoltStress * (1 + (CrossLoad / 100M)) * TensileStressArea_mm * LTFValue / 1000M), 2, MidpointRounding.AwayFromZero).ToString();
                    if (BoltYield > 0)
                    {
                        T1ABoltYield_SI_Label.Text = Math.Round(((ResidualBoltStress * (1 + (CrossLoad / 100M)) * LTFValue) * 100M / BoltYield), 2, MidpointRounding.AwayFromZero).ToString();
                    }
                    else
                    {
                        T1ABoltYield_SI_Label.Text = "";
                    }
                    TensioningPressureB = (ResidualBoltStress * TensileStressArea_mm * LTFValue / (ToolPressureArea));
                    TensioningPressureA = TensioningPressureB * (1 + (CrossLoad / 100M));
                    PressureA_In_Label.Text = Math.Round((TensioningPressureA * 145.037738007M), 0, MidpointRounding.AwayFromZero).ToString() + " psi";
                    PressureA_SI_Label.Text = Math.Round((TensioningPressureA), 2, MidpointRounding.AwayFromZero).ToString() + " MPa";
                    Pass1AppliedPressure_SequenceTabValue_In_Label.Text = Math.Round((TensioningPressureA * 145.037738007M), 0, MidpointRounding.AwayFromZero).ToString();
                    Pass1AppliedPressure_SequenceTabValue_SI_Label.Text = Math.Round(TensioningPressureA, 2, MidpointRounding.AwayFromZero).ToString();

                    // Warnings that permitted values are exceeded
                    if ((ResidualBoltStress * LTFValue * (1 + (CrossLoad / 100M))) > (BoltYield * 0.95M))     // Only 95% of bolt yield value, pressure A & B allowed
                    {
                        TitleWarning_PictureBox.Visible = true;
                        PressureAWarning_PictureBox.Visible = true;
                    }
                    if ((ResidualBoltStress * LTFValue) > (BoltYield * 0.95M))     // Only 95% of bolt yield value, pressure A & B allowed
                    {
                        TitleWarning_PictureBox.Visible = true;
                        PressureBWarning_PictureBox.Visible = true;
                    }
                    if ((ResidualBoltStress) > (BoltYield * 0.95M))
                    {
                        TitleWarning_PictureBox.Visible = true;
                        ResidualStreeWarning_PictureBox.Visible = true;
                    }
                    if (TensioningPressureA > 150)                           // Only 150 MPa tensioning is allowed.
                    {
                        TitleWarning_PictureBox.Visible = true;
                        PressureA_Info_Warning_PictureBox.Visible = true;
                    }
                    if (TensioningPressureB > 150)                           // Only 150 MPa tensioning is allowed.
                    {
                        TitleWarning_PictureBox.Visible = true;
                        PressureB_Info_Warning_PictureBox.Visible = true;
                    }
                    
                    break;

                case 2:
                    T1ABoltStress_In_Label.Text = Math.Round((ResidualBoltStress * LTFValue * (1 + (CrossLoad / 100M)) * 145.037738007M), 0, MidpointRounding.AwayFromZero).ToString(); ;
                    T1ABoltStress_SI_Label.Text = Math.Round((ResidualBoltStress * LTFValue * (1 + (CrossLoad / 100M))), 2, MidpointRounding.AwayFromZero).ToString(); ;
                    T1ABoltLoad_In_Label.Text = Math.Round((ResidualBoltStress * LTFValue * (1 + (CrossLoad / 100M)) * TensileStressArea_in * 145.037738007M / 2240M), 2, MidpointRounding.AwayFromZero).ToString();
                    T1ABoltLoad_SI_Label.Text = Math.Round((ResidualBoltStress * (1 + (CrossLoad / 100M)) * TensileStressArea_mm * LTFValue), 2, MidpointRounding.AwayFromZero).ToString();
                    if (BoltYield > 0)
                    {
                        T1ABoltYield_SI_Label.Text = Math.Round(((ResidualBoltStress * (1 + (CrossLoad / 100M)) * LTFValue) * 100M / BoltYield), 2, MidpointRounding.AwayFromZero).ToString();
                    }
                    else
                    {
                        T1ABoltYield_SI_Label.Text = "";
                    }
                    TensioningPressureB = (ResidualBoltStress * TensileStressArea_mm * LTFValue / (ToolPressureArea));
                    TensioningPressureA = TensioningPressureB * (1 + (CrossLoad / 100M));
                    PressureA_In_Label.Text = Math.Round((TensioningPressureA * 145.037738007M), 0, MidpointRounding.AwayFromZero).ToString() + " psi";
                    PressureA_SI_Label.Text = Math.Round((TensioningPressureA), 2, MidpointRounding.AwayFromZero).ToString() + " MPa";
                    Pass1AppliedPressure_SequenceTabValue_In_Label.Text = Math.Round((TensioningPressureA * 145.037738007M), 0, MidpointRounding.AwayFromZero).ToString();
                    Pass1AppliedPressure_SequenceTabValue_SI_Label.Text = Math.Round(TensioningPressureA, 2, MidpointRounding.AwayFromZero).ToString();
                    Pass2Bolt_SequenceTabValue_Label.Text = "3";
                    Pass2AppliedPressure_SequenceTabValue_In_Label.Text = Math.Round((TensioningPressureA * 145.037738007M), 0, MidpointRounding.AwayFromZero).ToString();
                    Pass2AppliedPressure_SequenceTabValue_SI_Label.Text = Math.Round(TensioningPressureA, 2, MidpointRounding.AwayFromZero).ToString();

                    // Warnings that permitted values are exceeded
                    if ((ResidualBoltStress * LTFValue * (1 + (CrossLoad / 100M))) > (BoltYield * 0.95M))     // Only 95% of bolt yield value, pressure A & B allowed
                    {
                        TitleWarning_PictureBox.Visible = true;
                        PressureAWarning_PictureBox.Visible = true;
                    }
                    if ((ResidualBoltStress * LTFValue) > (BoltYield * 0.95M))     // Only 95% of bolt yield value, pressure A & B allowed
                    {
                        TitleWarning_PictureBox.Visible = true;
                        PressureBWarning_PictureBox.Visible = true;
                    }
                    if ((ResidualBoltStress) > (BoltYield * 0.95M))
                    {
                        TitleWarning_PictureBox.Visible = true;
                        ResidualStreeWarning_PictureBox.Visible = true;
                    }
                    if (TensioningPressureA > 150)                           // Only 150 MPa tensioning is allowed.
                    {
                        TitleWarning_PictureBox.Visible = true;
                        PressureA_Info_Warning_PictureBox.Visible = true;
                    }
                    if (TensioningPressureB > 150)                           // Only 150 MPa tensioning is allowed.
                    {
                        TitleWarning_PictureBox.Visible = true;
                        PressureB_Info_Warning_PictureBox.Visible = true;
                    }
                    
                    break;
                case 3:
                    break;
            }

            // Graph
            DrawGraph();
        }

        private void PressureAWarning_PictureBox_MouseHover(object sender, EventArgs e)
        {
            MessageBox.Show("Bolt stress exceeds 95% of yield.");
        }

        private void PressureBWarning_PictureBox_MouseHover(object sender, EventArgs e)
        {
            MessageBox.Show("Bolt stress exceeds 95% of yield.");
        }

        private void ResidualStreeWarning_PictureBox_MouseHover(object sender, EventArgs e)
        {
            MessageBox.Show("Bolt stress exceeds 95% of yield.");
        }

        private void PressureA_Info_Warning_PictureBox_MouseHover(object sender, EventArgs e)
        {
            MessageBox.Show("Exceeds maximum working pressure of tensioning tool.");
        }

        private void PressureB_Info_Warning_PictureBox_MouseHover(object sender, EventArgs e)
        {
            MessageBox.Show("Exceeds maximum working pressure of tensioning tool.");
        }

        private void Deten_Info_Warning_PictureBox_MouseHover(object sender, EventArgs e)
        {
            MessageBox.Show("Exceeds maximum working pressure of tensioning tool.");
        }

        private void DefaultStress_Button_Click(object sender, EventArgs e)
        {
            //SQLiteConnection Stress_conn = new SQLiteConnection("Data Source=" + BasePath + "\\WizBolt.db;New=False;");
            //SQLiteCommand Stress_SQLiteCommand = new SQLiteCommand(Stress_conn);
            //Stress_conn.Open();
            //string RatingIdentity = Rating_ComboBox.SelectedValue.ToString();
            //string FlangeIdentity = Flange1Config_ComboBox.SelectedValue.ToString();

            //Stress_SQLiteCommand.CommandText = SELECT StandardId, StandardAuxiliaryId, FlangeRatingId, FlangeTypeId, ResidualBoltStress FROM Flanges 
            //     WHERE            StandardId = 
            //                  AND    StandardAuxiliaryId = 
            //                  AND    FlangeRatingId = 
            //                  AND    FlangeTypeId = 
        }

        private void UnitSystem_Button_Click(object sender, EventArgs e)
        {
            if (UnitSystem_Button.Text == "SAE")
            {
                UnitSystem_Button.Text = "S. I. Units";
                UnitSystem_Button.BorderColor = Color.ForestGreen;
                UnitSystem_Button.ColorFillBlend.iColor[0] = Color.GreenYellow;
                UnitSystem_Button.ColorFillBlend.iColor[1] = Color.ForestGreen;
                UnitSystem_Button.ColorFillBlend.iColor[2] = Color.DarkGreen;
                Clamp1Unit_Label.Text = "mm";
                GasketGapUnit_Label.Text = "mm";
                Clamp2Unit_Label.Text = "mm";
                CustomSpacerUnit_Label.Text = "mm";
                ResidualBoltStressUnit_Label.Text = "MPa";
                if (Clamp1_TextBox.Text.Length > 0)
                {
                    decimal TopClamp = Math.Round((Convert.ToDecimal(Clamp1_TextBox.Text) * 25.4M), 3, MidpointRounding.AwayFromZero);      // Convert to mm from inch
                    Clamp1_TextBox.Text = TopClamp.ToString();
                }
                if (Clamp2_TextBox.Text.Length > 0)
                {
                    decimal BottomClamp = Math.Round((Convert.ToDecimal(Clamp2_TextBox.Text) * 25.4M), 3, MidpointRounding.AwayFromZero);   // Convert to mm from inch
                    Clamp2_TextBox.Text = BottomClamp.ToString();
                }
                if (GasketGap_TextBox.Text.Length > 0)
                {
                    decimal GasketThickness = Math.Round((Convert.ToDecimal(GasketGap_TextBox.Text) * 25.4M), 3, MidpointRounding.AwayFromZero);
                    GasketGap_TextBox.Text = GasketThickness.ToString();
                }
                if (CustomSpacer_TextBox.Text.Length > 0)
                {
                    decimal CustomSpacer = Math.Round((Convert.ToDecimal(CustomSpacer_TextBox.Text) * 25.4M), 3, MidpointRounding.AwayFromZero);
                    CustomSpacer_TextBox.Text = CustomSpacer.ToString();
                }
                if (ResidualBoltStress_TextBox.Text.Length > 0)
                {
                    decimal ResidualBoltStress = Math.Round((Convert.ToDecimal(ResidualBoltStress_TextBox.Text) / 145.037738M), 3, MidpointRounding.AwayFromZero);
                    ResidualBoltStress_TextBox.Text = ResidualBoltStress.ToString();
                }
            }
            else
            {
                UnitSystem_Button.Text = "SAE";
                UnitSystem_Button.BorderColor = Color.SteelBlue;
                UnitSystem_Button.ColorFillBlend.iColor[0] = Color.AliceBlue;
                UnitSystem_Button.ColorFillBlend.iColor[1] = Color.RoyalBlue;
                UnitSystem_Button.ColorFillBlend.iColor[2] = Color.Navy;
                Clamp1Unit_Label.Text = "in";
                GasketGapUnit_Label.Text = "in";
                Clamp2Unit_Label.Text = "in";
                CustomSpacerUnit_Label.Text = "in";
                ResidualBoltStressUnit_Label.Text = "psi";
                if (Clamp1_TextBox.Text.Length > 0)
                {
                    decimal TopClamp = Math.Round((Convert.ToDecimal(Clamp1_TextBox.Text) / 25.4M), 4, MidpointRounding.AwayFromZero);      // Convert to inch from mm
                    Clamp1_TextBox.Text = TopClamp.ToString();
                }
                if (Clamp2_TextBox.Text.Length > 0)
                {
                    decimal BottomClamp = Math.Round((Convert.ToDecimal(Clamp2_TextBox.Text) / 25.4M), 4, MidpointRounding.AwayFromZero);   // Convert to inch from mm
                    Clamp2_TextBox.Text = BottomClamp.ToString();
                }
                if (GasketGap_TextBox.Text.Length > 0)
                {
                    decimal GasketThickness = Math.Round((Convert.ToDecimal(GasketGap_TextBox.Text) / 25.4M), 4, MidpointRounding.AwayFromZero);
                    GasketGap_TextBox.Text = GasketThickness.ToString();
                }
                if (CustomSpacer_TextBox.Text.Length > 0)
                {
                    decimal CustomSpacer = Math.Round((Convert.ToDecimal(CustomSpacer_TextBox.Text) / 25.4M), 4, MidpointRounding.AwayFromZero);
                    CustomSpacer_TextBox.Text = CustomSpacer.ToString();
                }
                if (ResidualBoltStress_TextBox.Text.Length > 0)
                {
                    decimal ResidualBoltStress = Math.Round((Convert.ToDecimal(ResidualBoltStress_TextBox.Text) * 145.037738M), 0, MidpointRounding.AwayFromZero);
                    ResidualBoltStress_TextBox.Text = ResidualBoltStress.ToString();
                }
            }
        }
       
   }
}
