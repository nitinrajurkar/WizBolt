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

namespace WizBolt
{
    public partial class Project : Form
    {
        public Project()
        {
            InitializeComponent();
        }

        public static WizBoltMainFrame WizMain = new WizBoltMainFrame();
        public int CurrentProjectId = 0;   // This is project report id along with project id as customer name and project name are merged together
        public int CustomerId = 0;         // Customer as well as project tables are separate as these are separate entities
        public int Number_of_ToolSeries = 0;    // Number of tool series available. Set after retrieving tool series from database table TensionerTools_Series
        public string BasePath = System.IO.Path.GetDirectoryName(Application.ExecutablePath);

        private void Cancel_Button_Click(object sender, EventArgs e)
        {
            this.Close();
            DialogResult = DialogResult.Cancel;
        }

        private void Project_Load(object sender, EventArgs e)
        {
            SQLiteConnection connect = new SQLiteConnection("Data Source=" + BasePath + "\\WizBolt.db;New=False;");
            SQLiteCommand ProjectPopulate_SQLiteCommand = new SQLiteCommand(connect);
            connect.Open();

            // Tensioner Tool Series: Gives Applied Area e.g. Topside (on ground), Subsea (under water), etc.
            try
            {
                ProjectPopulate_SQLiteCommand.CommandText = "SELECT TensionerToolSeriesId,  (COALESCE(AppliedField, ' ') || ' ' || COALESCE(Series, '')) AS ToolSeries FROM TensionerTools_Series";
                DataSet ToolRange_DataSet = new DataSet();
                SQLiteDataAdapter ToolRange_DataAdapter = new SQLiteDataAdapter(ProjectPopulate_SQLiteCommand);
                int ToolSeriesResult = ToolRange_DataAdapter.Fill(ToolRange_DataSet);
                CheckBox ToolSeries_CheckBox;
                if (ToolSeriesResult >= 0)
                {
                    //ToolRange_ComboBox.DataSource = ToolRange_DataSet.Tables[0];
                    //ToolRange_ComboBox.DisplayMember = "ToolSeries";
                    //ToolRange_ComboBox.ValueMember = "TensionerToolSeriesId";
                    Number_of_ToolSeries = ToolRange_DataSet.Tables[0].Rows.Count;
                    int Location_X = 154;
                    int Location_A = 154;
                    int TurnValue = 0;
                    for (int i = 0; i < Number_of_ToolSeries; i++)
                    {
                        ToolSeries_CheckBox = new CheckBox();
                        ToolSeries_CheckBox.Name = "ToolSeries" + i.ToString();         // To find the control
                        ToolSeries_CheckBox.Text = ToolRange_DataSet.Tables[0].Rows[i]["ToolSeries"].ToString(); // Display to user
                        ToolSeries_CheckBox.Tag = ToolRange_DataSet.Tables[0].Rows[i]["TensionerToolSeriesId"]; // To get the id of the tool
                        ToolSeries_CheckBox.AutoSize = true;
                        //ToolSeries_CheckBox.Top = 384;
                        //ToolSeries_CheckBox.Left = 160;
                        if (i > 0)
                        {
                            Location_X += (ToolRange_DataSet.Tables[0].Rows[i - 1]["ToolSeries"].ToString().Length + 36);
                            if (((i * 88) + Location_X + (ToolRange_DataSet.Tables[0].Rows[i]["ToolSeries"].ToString().Length) + (ToolRange_DataSet.Tables[0].Rows[0]["ToolSeries"].ToString().Length) + 36) < 510)
                            {
                                ToolSeries_CheckBox.Location = new Point(i * 88 + Location_X, 384);       // Horizontal First Line
                                TurnValue = i;
                            }
                            else if (((((i - (TurnValue + 1)) * 88) + Location_A + (ToolRange_DataSet.Tables[0].Rows[i]["ToolSeries"].ToString().Length) + (ToolRange_DataSet.Tables[0].Rows[TurnValue + 1]["ToolSeries"].ToString().Length) + 36) < 510))
                            {
                                if (i > (TurnValue + 1))
                                {
                                    Location_A += (ToolRange_DataSet.Tables[0].Rows[i - 1]["ToolSeries"].ToString().Length + 36);
                                    ToolSeries_CheckBox.Location = new Point((i - (TurnValue + 1)) * 88 + Location_A, 412);       // Horizontal Second Line
                                }
                                else
                                {
                                    ToolSeries_CheckBox.Location = new Point(Location_A, 412);
                                }
                            }
                            else
                            {
                                MessageBox.Show("Too many tool series. All tool series cannot be displayed here.");
                            }
                        }
                        else
                        {
                            ToolSeries_CheckBox.Location = new Point(Location_X, 384); 
                        }
                        this.Controls.Add(ToolSeries_CheckBox);
                    }
                }
            }
            catch (Exception ToolRange_Exception)
            {
                MessageBox.Show(ToolRange_Exception.Message + " while populating tool range in project.");
            }
        }

        private void Ok_Button_Click(object sender, EventArgs e)
        {
            Control[] InfoControls = WizBoltMainFrame.MainForm.Controls.Find("CoefficientValue_Label", true);
            InfoControls[0].Text = FrictionCoefficient_TextBox.Text;
            Control[] TabControls = WizBoltMainFrame.MainForm.Controls.Find("Coeff_TextBox", true);
            TabControls[0].Text = FrictionCoefficient_TextBox.Text;
            string ToolSeriesId = string.Empty;
            string ToolSeries = string.Empty;
            string StressBase = string.Empty;
            DateTime ProjectStartDate = Convert.ToDateTime(ProjectDate_DateTimePicker.Text);

            if (TensileStressArea_RadioButton.Checked)
            {
                StressBase = "Tensile Stress Area";
            }
            else
            {
                StressBase = "Minor Diameter Area";
            }

            for (int ToolId = 0; ToolId < Number_of_ToolSeries; ToolId++)
            {
                string Tool_CheckBoxName = "ToolSeries" + ToolId.ToString();         // To find the control with its name
                foreach (Control control in this.Controls)
                {
                    if ((control.Name == Tool_CheckBoxName) && (control is CheckBox))
                    {
                        CheckBox checkBox = control as CheckBox;
                        if (checkBox.Checked)
                        {
                            if ((ToolSeriesId.Length == 0) && (ToolSeriesId == string.Empty))
                            {
                                ToolSeriesId = checkBox.Tag.ToString();
                                ToolSeries = checkBox.Text;
                            }
                            else
                            {
                                ToolSeriesId = ToolSeriesId + " & " + checkBox.Tag.ToString();
                                ToolSeries = ToolSeries + " & " + checkBox.Text;
                            }
                        }
                    }
                }
            }
            if (ToolSeriesId == string.Empty)
            {
                MessageBox.Show("Please select at least one tensioner tool series.", "User Information!");
            }
            else
            {

                foreach (Control c in WizBoltMainFrame.MainForm.Controls)
                {
                    if (c.Name == "CrossLoading_TextBox")
                    {
                        c.Text = CrossLoading_TextBox.Text;
                    }

                    if (c.Name == "Detensioning_TextBox")
                    {
                        c.Text = Detensioning_TextBox.Text;
                    }

                    if (c.Name == "ToolSeriesId_Label")
                    {
                        c.Text = ToolSeriesId;
                        c.Visible = false;
                    }
                    
                    if (c.Name == "Notes_RichTextBox")
                    {
                        c.Text = Notes_RichTextBox.Text;
                        c.Visible = false;
                    }
                    if (c.Name == "Summary_Label")
                    {
                        c.Text = SummaryNotes_TextBox.Text;
                        c.Visible = false;
                    }
                    if (c.Name == "Stress_Label")
                    {
                        c.Text = StressBase;
                        c.Visible = false;
                    }
                   
                    if (c.Name == "Panel_Grouper")
                    {
                        foreach (Control ci in c.Controls)
                        {
                            if (ci.Name == "CustomerName_Label")
                            {
                                ci.Text = Client_TextBox.Text;
                                ci.Visible = true;
                            }

                            if (ci.Name == "Location_Label")
                            {
                                ci.Visible = true;
                            }

                            if (ci.Name == "LocationName_Label")
                            {
                                ci.Text = Location_TextBox.Text;
                                ci.Visible = true;
                            }

                            if (ci.Name == "ProjectName_Label")
                            {
                                ci.Text = Project_TextBox.Text;
                                ci.Visible = true;
                            }

                            if (ci.Name == "Reference_Label")
                            {
                                ci.Text = Reference_TextBox.Text;
                                ci.Visible = true;
                            }

                            if (ci.Name == "ProjectDate_Label")
                            {
                                ci.Visible = true;
                            }

                            if (ci.Name == "EnteredProjectDate_label")
                            {
                                ci.Text = ProjectDate_DateTimePicker.Value.ToShortDateString();
                                ci.Visible = true;
                            }
                            if (ci.Name == "ProjectEndDateTitle_Label")
                            {
                                ci.Visible = true;
                            }
                            if (ci.Name == "ProjectEndDate_Label")
                            {
                                ci.Visible = true;
                            }
                            if (ci.Name == "ToolSeries_Label")
                            {
                                ci.Text = ToolSeries;
                                ci.Visible = true;

                            }
                            if (ci.Name == "Engineer_Label")
                            {
                                ci.Text = Engineer_TextBox.Text;
                                ci.Visible = true;
                            }
                        }
                    }
                }

                if ((CrossLoading_TextBox.Text.Length == 0) && (CrossLoading_TextBox.Text == ""))
                {
                    MessageBox.Show("Please enter cross loading factor. If no cross loading then enter 0.", "Wrong Data Entered!");
                }
                else
                    if ((Detensioning_TextBox.Text.Length == 0) && (Detensioning_TextBox.Text == ""))
                    {
                        MessageBox.Show("Please enter detensioning factor. If detensioning then enter 0.", "Wrong Data Entered!");
                    }
                    else
                        if ((FrictionCoefficient_TextBox.Text.Length == 0) && (FrictionCoefficient_TextBox.Text == ""))
                        {
                            MessageBox.Show("Please enter coefficient of friction. ", "Wrong Data Entered!");
                        }
                        else
                        {
                            SQLiteConnection ConnectProject = new SQLiteConnection("Data Source=" + BasePath + "\\WizBolt.db;New=False;");
                            SQLiteCommand SaveProject_SQLiteCommand = new SQLiteCommand(ConnectProject);
                            ConnectProject.Open();
                            try
                            {
                                // Project information
                                SaveProject_SQLiteCommand.CommandText = "SELECT 	MAX(COALESCE(ProjectReportId, 0)) "
                                                                            + " FROM ProjectReport ";
                                SQLiteDataReader Project_DataReader = SaveProject_SQLiteCommand.ExecuteReader();
                                if (Project_DataReader.Read())
                                {
                                    if (Project_DataReader.IsDBNull(0))
                                    {
                                        ++CurrentProjectId;
                                    }
                                    else
                                    {
                                        CurrentProjectId = Project_DataReader.GetInt32(0) + 1;
                                    }
                                }
                                else
                                {
                                    ++CurrentProjectId;
                                }
                                if (!Project_DataReader.IsClosed)
                                {
                                    Project_DataReader.Close();
                                    Project_DataReader.Dispose();
                                }

                                // Customer Information
                                // If customer exists, then find it
                                SaveProject_SQLiteCommand.CommandText = "SELECT CustomerId, CustomerName  FROM Customer "
                                                                        + " WHERE CustomerName = '" + Client_TextBox.Text.Trim() + "'";

                                SQLiteDataReader Customer_DataReader = SaveProject_SQLiteCommand.ExecuteReader();
                                if (Customer_DataReader.Read())
                                {
                                    if (Customer_DataReader.IsDBNull(0))
                                    {
                                        SaveProject_SQLiteCommand.CommandText = "SELECT 	MAX(COALESCE(CustomerId, 0)) "
                                                                            + " FROM Customer ";
                                        SQLiteDataReader NewCustomer_DataReader = SaveProject_SQLiteCommand.ExecuteReader();
                                        if (NewCustomer_DataReader.Read())
                                        {
                                            if (NewCustomer_DataReader.IsDBNull(0))
                                            {
                                                ++CustomerId;
                                            }
                                            else
                                            {
                                                CustomerId = NewCustomer_DataReader.GetInt32(0) + 1;
                                            }
                                            NewCustomer_DataReader.Close();
                                            NewCustomer_DataReader.Dispose();
                                        }
                                        else
                                        {
                                            ++CustomerId;
                                            if (!NewCustomer_DataReader.IsClosed)
                                            {
                                                NewCustomer_DataReader.Close();
                                            }
                                        }

                                        SaveProject_SQLiteCommand.CommandText = "INSERT INTO Customer (CustomerId, CustomerName, CustomerAddress)"
                                                                                + " VALUES (" + CustomerId + ", '" + Client_TextBox.Text + "', " + Location_TextBox.Text + " );";

                                        int SaveCustomer = SaveProject_SQLiteCommand.ExecuteNonQuery();
                                        if (SaveCustomer > 0)
                                        {
                                            Control[] CustomerControls = WizBoltMainFrame.MainForm.Controls.Find("CustomerId_Label", true);
                                            CustomerControls[0].Text = CustomerId.ToString();
                                        }
                                    }
                                    else
                                    {
                                        CustomerId = Customer_DataReader.GetInt32(0);
                                    }
                                    Customer_DataReader.Close();
                                }
                                else
                                {
                                    if (!Customer_DataReader.IsClosed)
                                    {
                                        Customer_DataReader.Close();
                                    }

                                    SaveProject_SQLiteCommand.CommandText = "SELECT 	MAX(COALESCE(CustomerId, 0)) "
                                                                            + " FROM Customer ";
                                    SQLiteDataReader NewCustomer_DataReader = SaveProject_SQLiteCommand.ExecuteReader();
                                    if (NewCustomer_DataReader.Read())
                                    {
                                        if (NewCustomer_DataReader.IsDBNull(0))
                                        {
                                            ++CustomerId;
                                        }
                                        else
                                        {
                                            CustomerId = NewCustomer_DataReader.GetInt32(0) + 1;
                                        }
                                        NewCustomer_DataReader.Close();
                                        NewCustomer_DataReader.Dispose();
                                    }
                                    else
                                    {
                                        ++CustomerId;
                                        if (!NewCustomer_DataReader.IsClosed)
                                        {
                                            NewCustomer_DataReader.Close();
                                        }
                                    }

                                    SaveProject_SQLiteCommand.CommandText = "INSERT INTO Customer (CustomerId, CustomerName, CustomerAddress)"
                                                                            + " VALUES (" + CustomerId + ", '" + Client_TextBox.Text + "', '" + Location_TextBox.Text + "' );";

                                    int SaveCustomer = SaveProject_SQLiteCommand.ExecuteNonQuery();

                                }

                                if (CurrentProjectId > 0)
                                {
                                    // Save new project
                                    SaveProject_SQLiteCommand.CommandText = "INSERT INTO Projects (CustomerId, ProjectId, ProjectName, "
                                                                           + " ProjectReference, Description, StartDateDay, StartDateMonth, StartDateYear, "
                                                                           + "  Notes, SummaryNotes)"
                                                                           + " VALUES (" + CustomerId + ", " + CurrentProjectId + ", '" + Project_TextBox.Text + "', "
                                                                           + "'" + Reference_TextBox.Text + "', ' ', '" + ProjectStartDate.Day.ToString() + "', '" + ProjectStartDate.Month.ToString() + "', '" + ProjectStartDate.Year.ToString() + "', "
                                                                           + "'" + Notes_RichTextBox.Text + "', '" + SummaryNotes_TextBox.Text + "');";

                                    int SaveProject = SaveProject_SQLiteCommand.ExecuteNonQuery();

                                    // Save project report
                                    SaveProject_SQLiteCommand.CommandText = "INSERT INTO ProjectReport (ProjectReportId, CustomerId, CustomerName, CustomerLocation, ProjectName, "
                                                                           + " ProjectReference, StartDateDay, StartDateMonth, StartDateYear, "
                                                                           + " EngineerName,  Notes, SummaryNotes,"
                                                                           + " TensionerToolSeriesId, Series, CrossLoading_PC, Detensioning_PC, Coefficient_of_Friction, StressValue_Base"
                                                                           + ")"
                                                                           + " VALUES (" + CurrentProjectId + ", " + CustomerId + ", '" + Client_TextBox.Text + "',  '" + Location_TextBox.Text + "', '" + Project_TextBox.Text + "', "
                                                                           + "'" + Reference_TextBox.Text + "', '" + ProjectStartDate.Day.ToString() + "', '" + ProjectStartDate.Month.ToString() + "', '" + ProjectStartDate.Year.ToString() + "', "
                                                                           + "'" + Engineer_TextBox.Text + "', '" + Notes_RichTextBox.Text + "', '" + SummaryNotes_TextBox.Text + "', "
                                                                           + "'" + ToolSeriesId + "', '" + ToolSeries + "', " + CrossLoading_TextBox.Text + "," + Detensioning_TextBox.Text + ", " + FrictionCoefficient_TextBox.Text + ", '" + StressBase + "'"
                                                                           + ");";

                                    int SaveReport = SaveProject_SQLiteCommand.ExecuteNonQuery();
                                    if (SaveReport > 0)
                                    {
                                        Control[] ProjectControls = WizBoltMainFrame.MainForm.Controls.Find("ProjectId_Label", true);
                                        ProjectControls[0].Text = CurrentProjectId.ToString();
                                        MessageBox.Show("Project successfully saved. Please note that New Project alwasys create new project even if it has been created earlier. To continue with earlier project, please load or open it.", "Information to user");
                                    }
                                }
                            }
                            catch (Exception ProjectSave_Exception)
                            {
                                MessageBox.Show(ProjectSave_Exception.Message + " while saving new project.");
                            }
                            finally
                            {
                                ConnectProject.Close();
                                SaveProject_SQLiteCommand.Dispose();
                                ConnectProject.Dispose();
                            }
                            this.Close();
                            DialogResult = DialogResult.OK;
                        }
            }
        }
    }
}
