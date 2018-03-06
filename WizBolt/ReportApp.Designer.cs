namespace WizBolt
{
    partial class ReportApp
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            Microsoft.Reporting.WinForms.ReportDataSource reportDataSource1 = new Microsoft.Reporting.WinForms.ReportDataSource();
            this.PipeJointDataBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.BoltProjectData = new WizBolt.BoltProjectData();
            this.Application_ReportViewer = new Microsoft.Reporting.WinForms.ReportViewer();
            this.PipeJointDataTableAdapter = new WizBolt.BoltProjectDataTableAdapters.PipeJointDataTableAdapter();
            ((System.ComponentModel.ISupportInitialize)(this.PipeJointDataBindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.BoltProjectData)).BeginInit();
            this.SuspendLayout();
            // 
            // PipeJointDataBindingSource
            // 
            this.PipeJointDataBindingSource.DataMember = "PipeJointData";
            this.PipeJointDataBindingSource.DataSource = this.BoltProjectData;
            // 
            // BoltProjectData
            // 
            this.BoltProjectData.DataSetName = "BoltProjectData";
            this.BoltProjectData.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
            // 
            // Application_ReportViewer
            // 
            this.Application_ReportViewer.AutoSize = true;
            reportDataSource1.Name = "PipeJoint_DataSet";
            reportDataSource1.Value = this.PipeJointDataBindingSource;
            this.Application_ReportViewer.LocalReport.DataSources.Add(reportDataSource1);
            this.Application_ReportViewer.LocalReport.ReportEmbeddedResource = "WizBolt.PipeJointReport.rdlc";
            this.Application_ReportViewer.Location = new System.Drawing.Point(1, -1);
            this.Application_ReportViewer.Name = "Application_ReportViewer";
            this.Application_ReportViewer.ShowBackButton = false;
            this.Application_ReportViewer.ShowContextMenu = false;
            this.Application_ReportViewer.ShowCredentialPrompts = false;
            this.Application_ReportViewer.ShowDocumentMapButton = false;
            this.Application_ReportViewer.ShowFindControls = false;
            this.Application_ReportViewer.ShowPageNavigationControls = false;
            this.Application_ReportViewer.ShowParameterPrompts = false;
            this.Application_ReportViewer.ShowProgress = false;
            this.Application_ReportViewer.ShowPromptAreaButton = false;
            this.Application_ReportViewer.ShowRefreshButton = false;
            this.Application_ReportViewer.ShowStopButton = false;
            this.Application_ReportViewer.ShowZoomControl = false;
            this.Application_ReportViewer.Size = new System.Drawing.Size(772, 682);
            this.Application_ReportViewer.TabIndex = 0;
            this.Application_ReportViewer.Visible = false;
            // 
            // PipeJointDataTableAdapter
            // 
            this.PipeJointDataTableAdapter.ClearBeforeFill = true;
            // 
            // ReportApp
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(778, 750);
            this.Controls.Add(this.Application_ReportViewer);
            this.Name = "ReportApp";
            this.Text = "Application Details Report";
            this.Load += new System.EventHandler(this.ReportApp_Load);
            ((System.ComponentModel.ISupportInitialize)(this.PipeJointDataBindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.BoltProjectData)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private Microsoft.Reporting.WinForms.ReportViewer Application_ReportViewer;
        private System.Windows.Forms.BindingSource PipeJointDataBindingSource;
        private BoltProjectData BoltProjectData;
        private BoltProjectDataTableAdapters.PipeJointDataTableAdapter PipeJointDataTableAdapter;
    }
}