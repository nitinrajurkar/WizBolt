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
using Microsoft.Reporting.WinForms;

namespace WizBolt
{
    public partial class ReportApp : Form
    {
        public int SelectedProjectId = 0;
        public int SelectedApplicationId = 0;
        public string BasePath = System.IO.Path.GetDirectoryName(Application.ExecutablePath);

        public ReportApp()
        {
            InitializeComponent();
        }

        private void ReportApp_Load(object sender, EventArgs e)
        {
            Control[] ProjectControl = WizBoltMainFrame.MainForm.Controls.Find("ProjectId_Label", true);
            SelectedProjectId = Convert.ToInt32((ProjectControl[0].Text));
            Control[] AppControl = WizBoltMainFrame.MainForm.Controls.Find("AppId_Label", true);
            SelectedApplicationId = Convert.ToInt32(AppControl[0].Text);

            // Following is very critical and important. The class here is the name of the DataSet that is created using new ... See the list in Solution Explorer.
            // It (i.e. BoltProjectData) has Table Adapter PipeJointDataTableAdapter which in turn have a table named PipeJointData. Mark these names where they are used carefully.
            // Most important is, for the report the dataset name given is PipeJoint_DataSet and this is where you assign new data.
            BoltProjectData MainAppData = new BoltProjectData();
            
            if (SelectedProjectId > 0)
            {
                MainAppData = RetrieveData();
            }
            else
            {
                MainAppData = RetrieveDefaultData();
            }

            // PipeJoint_DataSet is the one named as such when assigning table adapter to report. In Report Data Window you will see this name when you click on DataSet
            ReportDataSource App_ReportDataSource = new ReportDataSource("PipeJoint_DataSet", MainAppData.Tables["PipeJointData"]);   
            this.Application_ReportViewer.LocalReport.DataSources.Clear();
            this.Application_ReportViewer.LocalReport.DataSources.Add(App_ReportDataSource);
            this.Application_ReportViewer.Visible = true;
            this.Application_ReportViewer.RefreshReport();
            // This line of code loads data into the 'BoltProjectData.PipeJointData' table. 
            //this.PipeJointDataTableAdapter.Fill(this.BoltProjectData.PipeJointData);
            //this.Application_ReportViewer.RefreshReport();
        }

        private BoltProjectData RetrieveData()
        {
            BoltProjectData SelectedAppData = new BoltProjectData();           // There are three tables in this currently. They are PipeJoinData, ProjectReport, ApplicationReport.
            try
            {
                SQLiteConnection ConnectApp = new SQLiteConnection("Data Source=" + BasePath + "\\WizBolt.db;New=False;");
                SQLiteCommand PopulateApp_SQLiteCommand = new SQLiteCommand(ConnectApp);
                PopulateApp_SQLiteCommand.CommandText = "SELECT  PR.ProjectReportId,  PR.CustomerName, PR.ProjectName, PR.ProjectReference,  PR.StartDateDay, PR.StartDateMonth, PR.StartDateYear,    "
                                                        + "PR.EndDateDay, PR.EndDateMonth, PR.EndDateYear, PR.EngineerName,  "
                                                        + "PD.ApplicationId, PD.UnitSystem,  PD.JointId, PrimaryStandardId,  Specification, PrimaryFlangeRatingId, FlangeRating, Flange1_TypeId,  "
                                                        + " Flange1_Abbreviation, Flange1_ClampLength, Flange2_TypeId, Flange2_Abbreviation,  "
                                                        + " Flange2_ClampLength, GasketId, Gasket, GasketGap,  "
                                                        + " SpacerId, Spacer, SpacerThickness, TotalClampLength, PrimaryFlangeBoltId,  "
                                                        + " BoltThread, Pitch_or_TPI, Pitch_TPI_Value, Number_of_Bolts, LTF, "
                                                        + " Bolt_to_ToolRatioId, Bolt_to_ToolRatio, ToolId, ModelNumber,  "
                                                        + " ToolPressureArea, MaterialId, Material, BoltYield, BoltTensileStressArea, "
                                                        + " BoltLength, BoltStressBase, BoltMinorDiameterArea, T1_APressureBoltStress, T1_APressureBoltLoad,  "
                                                        + " T1_APressureBoltYieldPC, T1_BPressureBoltStress, T1_BPressureBoltLoad, T1_BPressureBoltYieldPC,  "
                                                        + " T2_ResidualBoltStress, T2_ResidualBoltLoad, T2_ResidualBoltYieldPC, DetensioningBoltStress, DetensioningBoltLoad,	"
                                                        + " DetensioningBoltYieldPC, TensionPressure_FirstPass, TensionPressure_SecondPass, TensionPressure_ThirdPass, CheckingPass,   "
                                                        + " PD.Torque, PD.Coefficient_Friction, PD.Bolt01, PD.Bolt02, PD.Bolt03, PD.Bolt04, "
                                                        + " PD.FirstPass_100, PD.FirstPass_50, PD.SecondPass_50, Max_Detensioning, ResidualBoltStress, Comments, IsDetension, CrossLoadingPC, "
                                                        + " PD.DetensioningPC, T3_ResidualBoltStress, T3_ResidualBoltLoad, T3_ResidualBoltYieldPC    "
                                                        + " FROM     ProjectReport       PR "
                                                        + "            INNER JOIN   ProjectDetailedReport        PD		ON		PD.ProjectReportId  =  PR.ProjectReportId "
                                                        + "                     WHERE  PR.ProjectReportId > 0 AND PR.ProjectReportId = " + SelectedProjectId.ToString();
                                                                                   //     + "  AND PD.ApplicationId = " + SelectedApplicationId.ToString() + ";";

                SQLiteDataAdapter AppReport_Adapter = new SQLiteDataAdapter(PopulateApp_SQLiteCommand);

                int ReportResult = AppReport_Adapter.Fill(SelectedAppData, "PipeJointData"); // There are currently three tables. So check the names in SelectedAppData. You may see first table blank. The data is in third table. In BoltProjectData the Table Adapter name is PipeJointDataTableAdapter and its table name is PipeJointData

                return SelectedAppData;
            }
            catch (Exception AppReport_Exception)
            {
                MessageBox.Show(AppReport_Exception.Message + " while displaying report.", "Error!");
                return SelectedAppData;
            }
        }

        private BoltProjectData RetrieveDefaultData()
        {
            BoltProjectData SelectedAppData = new BoltProjectData();
            try
            {
                SQLiteConnection ConnectApp = new SQLiteConnection("Data Source=" + BasePath + "\\WizBolt.db;New=False;");
                SQLiteCommand PopulateApp_SQLiteCommand = new SQLiteCommand(ConnectApp);
                PopulateApp_SQLiteCommand.CommandText = "SELECT  PR.ProjectReportId,  PR.CustomerName, PR.ProjectName, PR.ProjectReference,  PR.StartDateDay, PR.StartDateMonth, PR.StartDateYear,    "
                                                        + "PR.EndDateDay, PR.EndDateMonth, PR.EndDateYear, PR.EngineerName,  "
                                                        + "PD.ApplicationId, PD.UnitSystem,  PD.JointId, PrimaryStandardId,  Specification, PrimaryFlangeRatingId, FlangeRating, Flange1_TypeId,  "
                                                        + " Flange1_Abbreviation, Flange1_ClampLength, Flange2_TypeId, Flange2_Abbreviation,  "
                                                        + " Flange2_ClampLength, GasketId, Gasket, GasketGap,  "
                                                        + " SpacerId, Spacer, SpacerThickness, TotalClampLength, PrimaryFlangeBoltId,  "
                                                        + " BoltThread, Pitch_or_TPI, Pitch_TPI_Value, Number_of_Bolts, LTF, "
                                                        + " Bolt_to_ToolRatioId, Bolt_to_ToolRatio, ToolId, ModelNumber,  "
                                                        + " ToolPressureArea, MaterialId, Material, BoltYield, BoltTensileStressArea, "
                                                        + " BoltLength, BoltStressBase, BoltMinorDiameterArea, T1_APressureBoltStress, T1_APressureBoltLoad,  "
                                                        + " T1_APressureBoltYieldPC, T1_BPressureBoltStress, T1_BPressureBoltLoad, T1_BPressureBoltYieldPC,  "
                                                        + " T2_ResidualBoltStress, T2_ResidualBoltLoad, T2_ResidualBoltYieldPC, DetensioningBoltStress, DetensioningBoltLoad,	"
                                                        + " DetensioningBoltYieldPC, TensionPressure_FirstPass, TensionPressure_SecondPass, TensionPressure_ThirdPass, CheckingPass,   "
                                                        + " PD.Torque, PD.Coefficient_Friction, PD.Bolt01, PD.Bolt02, PD.Bolt03, PD.Bolt04, "
                                                        + " PD.FirstPass_100, PD.FirstPass_50, PD.SecondPass_50, Max_Detensioning, ResidualBoltStress, Comments, IsDetension, CrossLoadingPC, "
                                                        + " PD.DetensioningPC, T3_ResidualBoltStress, T3_ResidualBoltLoad, T3_ResidualBoltYieldPC   "
                                                        + " FROM     ProjectReport       PR "
                                                        + "            INNER JOIN   ProjectDetailedReport        PD		ON		PD.ProjectReportId  =  PR.ProjectReportId "
                                                        + "                     WHERE  PR.ProjectReportId = 0 AND PD.ApplicationId = 1; ";

                SQLiteDataAdapter AppReport_Adapter = new SQLiteDataAdapter(PopulateApp_SQLiteCommand);

                int ReportResult = AppReport_Adapter.Fill(SelectedAppData, "PipeJointData");        // In BoltProjectData the Table Adapter name is PipeJointDataTableAdapter and its table name is PipeJointData

                return SelectedAppData;
            }
            catch (Exception DefaultAppReport_Exception)
            {
                MessageBox.Show(DefaultAppReport_Exception.Message + " when default data for report uploaded.", "Error!");
                return SelectedAppData;
            }
        }
    }
}
