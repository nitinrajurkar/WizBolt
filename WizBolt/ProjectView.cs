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
using System.Globalization;
using System.Threading;

namespace WizBolt
{
    public partial class ProjectView : Form
    {
        public ProjectView()
        {
            InitializeComponent();
        }

        // Global Parameters
        #region Global Parameters
            Main MainForm = new Main();
            public string ProjectDisplayId = string.Empty;
            public string BasePath = System.IO.Path.GetDirectoryName(Application.ExecutablePath);
            public int Number_of_ToolSeries = 0;    // Number of tool series available. Set after retrieving tool series from database table TensionerTools_Series
            public string Selected_ToolSeriesId = string.Empty;  // To store the tool series that were selected during project creation
            public DateTime StartDate;
            public DateTime EndDate;
            public string Project_StartDateDay = string.Empty, 
                          Project_StartDateMonth = string.Empty, 
                          Project_StartDateYear = string.Empty,
                          Project_EndDateDay = string.Empty,
                          Project_EndDateMonth = string.Empty,
                          Project_EndDateYear = string.Empty;
            public string Date_Initial = string.Empty, Date_Middle = string.Empty, Date_End = string.Empty;      // Used to store single character of date format in capital letter.
            public string StartDate_RegionalFormat = string.Empty, EndDate_RegionalFormat = string.Empty;     // Date string as per the selected regional settings
        #endregion Global Parameters
        private void ProjectView_Load(object sender, EventArgs e)
        {
             foreach (Control c in WizBoltMainFrame.MainForm.Controls)
            {
                if (c.Name == "ToolSeriesId_Label")
                {
                    Selected_ToolSeriesId = c.Text;
                }
             }
            string[] Selected_ToolSeries_Identities = Selected_ToolSeriesId.Split('&');
            SQLiteConnection Project_Connection = new SQLiteConnection("Data Source=" + BasePath + "\\WizBolt.db;New=False;");
            SQLiteCommand Project_Command = new SQLiteCommand(Project_Connection);
            Project_Connection.Open();

            try
            {
                Project_Command.CommandText = "SELECT TensionerToolSeriesId,  (COALESCE(AppliedField, ' ') || ' ' || COALESCE(Series, '')) AS ToolSeries FROM TensionerTools_Series";
                DataSet ToolRange_DataSet = new DataSet();
                SQLiteDataAdapter ToolRange_DataAdapter = new SQLiteDataAdapter(Project_Command);
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
                                ToolSeries_CheckBox.Location = new Point(i * 88 + Location_X, 378);       // Horizontal First Line
                                TurnValue = i;
                            }
                            else if (((((i - (TurnValue + 1)) * 88) + Location_A + (ToolRange_DataSet.Tables[0].Rows[i]["ToolSeries"].ToString().Length) + (ToolRange_DataSet.Tables[0].Rows[TurnValue + 1]["ToolSeries"].ToString().Length) + 36) < 510))
                            {
                                if (i > (TurnValue + 1))
                                {
                                    Location_A += (ToolRange_DataSet.Tables[0].Rows[i - 1]["ToolSeries"].ToString().Length + 36);
                                    ToolSeries_CheckBox.Location = new Point((i - (TurnValue + 1)) * 88 + Location_A, 406);       // Horizontal Second Line
                                }
                                else
                                {
                                    ToolSeries_CheckBox.Location = new Point(Location_A, 406);
                                }
                            }
                            else
                            {
                                MessageBox.Show("Too many tool series. All tool series cannot be displayed here.");
                            }
                        }
                        else
                        {
                            ToolSeries_CheckBox.Location = new Point(Location_X, 378); 
                        }
                        this.Controls.Add(ToolSeries_CheckBox);
                    }
                
                }
            }
            catch (Exception ToolRange_Exception)
            {
                MessageBox.Show(ToolRange_Exception.Message + " while populating tool range in project view/modify.");
            }

            try
            {
                //Project_Command.CommandText = " SELECT  ProjectReportId,  CustomerName, CustomerLocation, ProjectName, ProjectReference,  "
                //                                + "  (StartDateDay || '/' || StartDateMonth || '/' || StartDateYear) AS StartDate,    "
                //                                + "  (EndDateDay || '/' || EndDateMonth || '/' || EndDateYear) AS EndDate,  EndDateDay,  "
                //                                + "  EngineerName,  Notes, SummaryNotes, TensionerToolSeriesId, Series,               "
                //                                + "  CrossLoading_PC, Detensioning_PC, Coefficient_of_Friction, StressValue_Base      "
                //                                + "  	FROM ProjectReport                                                            "
                //                                + "  			WHERE  ProjectReportId > 0 AND ProjectReportId = " + ProjectDisplayId;

                Project_Command.CommandText = " SELECT  ProjectReportId,  CustomerName, CustomerLocation, ProjectName, ProjectReference,  "
                                               + "  StartDateDay,  StartDateMonth, StartDateYear,    "
                                               + "  EndDateDay, EndDateMonth,  EndDateYear, "
                                               + "  EngineerName,  Notes, SummaryNotes, TensionerToolSeriesId, Series,               "
                                               + "  CrossLoading_PC, Detensioning_PC, Coefficient_of_Friction, StressValue_Base      "
                                               + "  	FROM ProjectReport                                                            "
                                               + "  			WHERE  ProjectReportId > 0 AND ProjectReportId = " + ProjectDisplayId;

                DataSet UpdateProject_DataSet = new DataSet();
                SQLiteDataAdapter UpdateProject_Adapter = new SQLiteDataAdapter(Project_Command);

                int Project_Result = UpdateProject_Adapter.Fill(UpdateProject_DataSet);
                if (Project_Result > 0)
                {
                    Client_TextBox.Text = UpdateProject_DataSet.Tables[0].Rows[0]["CustomerName"].ToString();
                    Location_TextBox.Text = UpdateProject_DataSet.Tables[0].Rows[0]["CustomerLocation"].ToString();
                    Project_TextBox.Text = UpdateProject_DataSet.Tables[0].Rows[0]["ProjectName"].ToString();
                    Reference_TextBox.Text = UpdateProject_DataSet.Tables[0].Rows[0]["ProjectReference"].ToString();
                    
                    //ProjectDate_DateTimePicker.Text
                    //EndDate_DateTimePicker.Text 
                    
                    Engineer_TextBox.Text = UpdateProject_DataSet.Tables[0].Rows[0]["EngineerName"].ToString();
                    Notes_RichTextBox.Text = UpdateProject_DataSet.Tables[0].Rows[0]["Notes"].ToString();
                    SummaryNotes_TextBox.Text = UpdateProject_DataSet.Tables[0].Rows[0]["SummaryNotes"].ToString();
                    //ToolRange_ComboBox.SelectedValue = UpdateProject_DataSet.Tables[0].Rows[0]["TensionerToolSeriesId"].ToString();
                    //ToolRange_ComboBox.SelectedItem = UpdateProject_DataSet.Tables[0].Rows[0]["Series"].ToString();
                    if (Selected_ToolSeries_Identities.Count() > 0)
                    {
                        for (int SeriesId = 0; SeriesId < Selected_ToolSeries_Identities.Count(); SeriesId++)
                        {
                            string Tool_CheckBoxTag = Selected_ToolSeries_Identities[SeriesId].Trim();         // To find the control with tag as tag is Tool Series Id
                            for (int ToolCount = 0; ToolCount < Number_of_ToolSeries; ToolCount++)
                            {
                                foreach (Control control in this.Controls)
                                {
                                    if (control is CheckBox)
                                    {
                                        string TagVal = control.Tag.ToString();
                                        if (TagVal == Tool_CheckBoxTag)
                                        {
                                            CheckBox checkBox = control as CheckBox;
                                            checkBox.Checked = true;
                                        }
                                    }
                                }
                            }
                        }
                    }
                    
                    CrossLoading_TextBox.Text = UpdateProject_DataSet.Tables[0].Rows[0]["CrossLoading_PC"].ToString();
                    Detensioning_TextBox.Text = UpdateProject_DataSet.Tables[0].Rows[0]["Detensioning_PC"].ToString();
                    FrictionCoefficient_TextBox.Text = UpdateProject_DataSet.Tables[0].Rows[0]["Coefficient_of_Friction"].ToString();
                    if (UpdateProject_DataSet.Tables[0].Rows[0]["StressValue_Base"].ToString().Trim() == "Tensile Stress Area")
                    {
                        TensileStressArea_RadioButton.Checked = true;
                        MinorDiameterArea_RadioButton.Checked = false;
                    }
                    else
                    {
                        TensileStressArea_RadioButton.Checked = false;
                        MinorDiameterArea_RadioButton.Checked = true;
                    }

                    Thread.CurrentThread.CurrentCulture.ClearCachedData();
                    var thread = new Thread(
                        s => ((DateCultures)s).Result = Thread.CurrentThread.CurrentCulture);
                    var CurrentCulturalSetting = new DateCultures();
                    thread.Start(CurrentCulturalSetting);
                    thread.Join();
                    var Culture = CurrentCulturalSetting.Result;
                     
                   

                   Project_StartDateDay = UpdateProject_DataSet.Tables[0].Rows[0]["StartDateDay"].ToString();
                   Project_StartDateMonth = UpdateProject_DataSet.Tables[0].Rows[0]["StartDateMonth"].ToString();
                   Project_StartDateYear = UpdateProject_DataSet.Tables[0].Rows[0]["StartDateYear"].ToString();
                    if ((UpdateProject_DataSet.Tables[0].Rows[0]["EndDateDay"].ToString() == "00") || (UpdateProject_DataSet.Tables[0].Rows[0]["EndDateDay"].ToString() == "0"))
                    {

                    }
                    else if ((!string.IsNullOrEmpty(UpdateProject_DataSet.Tables[0].Rows[0]["EndDateDay"].ToString())) && (!string.IsNullOrWhiteSpace(UpdateProject_DataSet.Tables[0].Rows[0]["EndDateDay"].ToString())))
                    {
                        Project_EndDateDay = UpdateProject_DataSet.Tables[0].Rows[0]["EndDateDay"].ToString();
                        Project_EndDateMonth = UpdateProject_DataSet.Tables[0].Rows[0]["EndDateMonth"].ToString();
                        Project_EndDateYear = UpdateProject_DataSet.Tables[0].Rows[0]["EndDateYear"].ToString();
                    }

                    DateTimeFormatInfo Current_DateTimeFormat = Culture.DateTimeFormat;
                    string Current_DateSeparator = Current_DateTimeFormat.DateSeparator;
                    string Current_ShortDatePattern = Current_DateTimeFormat.ShortDatePattern;
                    string DateFormat_Initial = Current_ShortDatePattern.Substring(0, Current_ShortDatePattern.IndexOf(Current_DateSeparator));
                    int Separator1 = Current_ShortDatePattern.IndexOf(Current_DateSeparator);
                    string DateRemainder = Current_ShortDatePattern.Substring(Separator1 + 1);
                    string DateFormat_Middle = DateRemainder.Substring(0, DateRemainder.IndexOf(Current_DateSeparator));
                    int Separator2 = DateRemainder.IndexOf(Current_DateSeparator);
                    string DateFormat_End = DateRemainder.Substring(Separator2 + 1);

                    Date_Initial = DateFormat_Initial.Substring(0, 1).ToUpper();
                    Date_Middle = DateFormat_Middle.Substring(0, 1).ToUpper();
                    Date_End = DateFormat_End.Substring(0, 1).ToUpper();

                    switch (Date_Initial)
                    {
                        case "D" :
                            if ((!string.IsNullOrEmpty(Project_StartDateDay)) && (!string.IsNullOrWhiteSpace(Project_StartDateDay)))
                            {
                                StartDate_RegionalFormat = Project_StartDateDay.Trim() + Current_DateSeparator.Trim();
                            }
                            
                            if ((!string.IsNullOrEmpty(Project_EndDateDay)) && (!string.IsNullOrWhiteSpace(Project_EndDateDay)))
                            {
                                EndDate_RegionalFormat = Project_EndDateDay.Trim() + Current_DateSeparator.Trim();
                            }
                            break;
                        case "M" :
                            if ((!string.IsNullOrEmpty(Project_StartDateMonth)) && (!string.IsNullOrWhiteSpace(Project_StartDateMonth)))
                            {
                                StartDate_RegionalFormat = Project_StartDateMonth.Trim() + Current_DateSeparator.Trim();
                            }
                            
                            if ((!string.IsNullOrEmpty(Project_EndDateMonth)) && (!string.IsNullOrWhiteSpace(Project_EndDateMonth)))
                            {
                                EndDate_RegionalFormat = Project_EndDateMonth.Trim() + Current_DateSeparator.Trim();
                            }
                            break;
                        case "Y" :
                            if ((!string.IsNullOrEmpty(Project_StartDateYear)) && (!string.IsNullOrWhiteSpace(Project_StartDateYear)))
                            {
                                StartDate_RegionalFormat = Project_StartDateYear.Trim() + Current_DateSeparator.Trim();
                            }
                            
                            if ((!string.IsNullOrEmpty(Project_EndDateYear)) && (!string.IsNullOrWhiteSpace(Project_EndDateYear)))
                            {
                                EndDate_RegionalFormat = Project_EndDateYear.Trim() + Current_DateSeparator.Trim();
                            }
                            break;
                        default :
                            MessageBox.Show("Date settings or date has encoutered a problem. Set up date format from 'Regional Settings in Control Panel' or check the entered date.");
                            break;
                    }

                    switch (Date_Middle)
                    {
                        case "D":
                            if ((!string.IsNullOrEmpty(Project_StartDateDay)) && (!string.IsNullOrWhiteSpace(Project_StartDateDay)))
                            {
                                StartDate_RegionalFormat = StartDate_RegionalFormat + Project_StartDateDay.Trim() + Current_DateSeparator.Trim();
                            }

                            if ((!string.IsNullOrEmpty(Project_EndDateDay)) && (!string.IsNullOrWhiteSpace(Project_EndDateDay)))
                            {
                                EndDate_RegionalFormat = EndDate_RegionalFormat + Project_EndDateDay.Trim() + Current_DateSeparator.Trim();
                            }
                            break;
                        case "M":
                            if ((!string.IsNullOrEmpty(Project_StartDateMonth)) && (!string.IsNullOrWhiteSpace(Project_StartDateMonth)))
                            {
                                StartDate_RegionalFormat = StartDate_RegionalFormat + Project_StartDateMonth.Trim() + Current_DateSeparator.Trim();
                            }

                            if ((!string.IsNullOrEmpty(Project_EndDateMonth)) && (!string.IsNullOrWhiteSpace(Project_EndDateMonth)))
                            {
                                EndDate_RegionalFormat = EndDate_RegionalFormat + Project_EndDateMonth.Trim() + Current_DateSeparator.Trim();
                            }
                            break;
                        case "Y":
                            if ((!string.IsNullOrEmpty(Project_StartDateYear)) && (!string.IsNullOrWhiteSpace(Project_StartDateYear)))
                            {
                                StartDate_RegionalFormat = StartDate_RegionalFormat + Project_StartDateYear.Trim() + Current_DateSeparator.Trim();
                            }

                            if ((!string.IsNullOrEmpty(Project_EndDateYear)) && (!string.IsNullOrWhiteSpace(Project_EndDateYear)))
                            {
                                EndDate_RegionalFormat = EndDate_RegionalFormat + Project_EndDateYear.Trim() + Current_DateSeparator.Trim();
                            }
                            break;
                        default:
                            MessageBox.Show("Date settings or date has encoutered a problem. Set up date format from 'Regional Settings in Control Panel' or check the entered date.");
                            break;
                    }

                    switch (Date_End)
                    {
                        case "D":
                            if ((!string.IsNullOrEmpty(Project_StartDateDay)) && (!string.IsNullOrWhiteSpace(Project_StartDateDay)))
                            {
                                StartDate_RegionalFormat = StartDate_RegionalFormat + Project_StartDateDay.Trim();
                            }

                            if ((!string.IsNullOrEmpty(Project_EndDateDay)) && (!string.IsNullOrWhiteSpace(Project_EndDateDay)))
                            {
                                EndDate_RegionalFormat = EndDate_RegionalFormat + Project_EndDateDay.Trim();
                            }
                            break;
                        case "M":
                            if ((!string.IsNullOrEmpty(Project_StartDateMonth)) && (!string.IsNullOrWhiteSpace(Project_StartDateMonth)))
                            {
                                StartDate_RegionalFormat = StartDate_RegionalFormat + Project_StartDateMonth.Trim();
                            }

                            if ((!string.IsNullOrEmpty(Project_EndDateMonth)) && (!string.IsNullOrWhiteSpace(Project_EndDateMonth)))
                            {
                                EndDate_RegionalFormat = EndDate_RegionalFormat + Project_EndDateMonth.Trim();
                            }
                            break;
                        case "Y":
                            if ((!string.IsNullOrEmpty(Project_StartDateYear)) && (!string.IsNullOrWhiteSpace(Project_StartDateYear)))
                            {
                                StartDate_RegionalFormat = StartDate_RegionalFormat + Project_StartDateYear.Trim();
                            }

                            if ((!string.IsNullOrEmpty(Project_EndDateYear)) && (!string.IsNullOrWhiteSpace(Project_EndDateYear)))
                            {
                                EndDate_RegionalFormat = EndDate_RegionalFormat + Project_EndDateYear.Trim();
                            }
                            break;
                        default:
                            MessageBox.Show("Date settings or date has encoutered a problem. Set up date format from 'Regional Settings in Control Panel' or check the entered date.");
                            break;
                    }

                    //if (DateTime.TryParseExact(Unconverted_StartDate, "dd/MM/yyyy h:mm:ss tt", CultureInfo.InvariantCulture, DateTimeStyles.None, out StartDate))
                    //{
                    //    ProjectDate_DateTimePicker.Value = StartDate;
                    //}
                    //else
                    //{

                    //}
                    ProjectDate_DateTimePicker.Text = StartDate_RegionalFormat;
                    if ((!string.IsNullOrEmpty(EndDate_RegionalFormat) && (!string.IsNullOrWhiteSpace(EndDate_RegionalFormat))))
                    {
                        EndDate_DateTimePicker.Text = EndDate_RegionalFormat;
                    }
                }
                else
                {
                    MessageBox.Show("Project not found. Either there may not be any current project or project is not saved. If there is no project, either upload existing project or create new", "Error!");
                }
            }
            catch (Exception UpdateProject_Exception)
            {
                MessageBox.Show(UpdateProject_Exception.Message + " while displaying/modifying project.", "Fatal Error!!!");
            }
            finally
            {
                Project_Connection.Close();
                Project_Command.Dispose();
                Project_Connection.Dispose();
            }
        }

        private void Cancel_Button_Click(object sender, EventArgs e)
        {
            this.Close();
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
            DateTime ProjectEndDate = Convert.ToDateTime(EndDate_DateTimePicker.Text);

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
                            ci.Text = EndDate_DateTimePicker.Value.ToShortDateString();
                            ci.Visible = true;
                        }
                        
                        if (ci.Name == "ToolSeries_Label")
                        {
                            ci.Text = ToolSeries;
                            ci.Visible = true;

                        }
                        if (ci.Name == "ToolSeriesId_Label")
                        {
                            ci.Text = ToolSeriesId;
                            ci.Visible = false;
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
                        SQLiteCommand UpdateProject_SQLiteCommand = new SQLiteCommand(ConnectProject);
                        ConnectProject.Open();
                        try
                        {
                            UpdateProject_SQLiteCommand.CommandText =  "   UPDATE ProjectReport SET ProjectName = '" + Project_TextBox.Text.Trim() + "', "
                                                                     + " CustomerName = '" + Client_TextBox.Text.Trim() + "', CustomerLocation = '" + Location_TextBox.Text + "', ProjectReference = '" + Reference_TextBox.Text.Trim() +"', "
                                                                     + " StartDateDay = " + ProjectStartDate.Day.ToString() + ", StartDateMonth = " + ProjectStartDate.Month.ToString() + ", StartDateYear = " + ProjectStartDate.Year.ToString()
                                                                     + ", EngineerName = '" + Engineer_TextBox.Text.Trim() + "', Notes = '" + Notes_RichTextBox.Text.Trim() + "', "
                                                                     + " SummaryNotes = '" + SummaryNotes_TextBox.Text.Trim() + "', TensionerToolSeriesId = '" + ToolSeriesId + "', Series = '" + ToolSeries + "', "
                                                                     + " CrossLoading_PC = " + CrossLoading_TextBox.Text.Trim() + ", Detensioning_PC = " + Detensioning_TextBox.Text.Trim() + ", "
                                                                     + " Coefficient_of_Friction = " + FrictionCoefficient_TextBox.Text.Trim() + ", StressValue_Base = '" + StressBase + "', "
                                                                     + " EndDateDay = " + ProjectEndDate.Day.ToString() + ", EndDateMonth = " + ProjectEndDate.Month.ToString() + ", EndDateYear = " + ProjectEndDate.Year.ToString()
				                                                     + "   WHERE ProjectReportId = " + ProjectDisplayId;
                            UpdateProject_SQLiteCommand.ExecuteNonQuery(); 
                            
                        }
                        catch (Exception ProjectModify_Exception)
                        {
                            MessageBox.Show(ProjectModify_Exception.Message + " while modifying the project.", "Fatal Error!!!");
                        }
                        finally
                        {
                            ConnectProject.Close();
                            UpdateProject_SQLiteCommand.Dispose();
                            ConnectProject.Dispose();
                        }
                        this.Close();
                    }
        }



    }
}
