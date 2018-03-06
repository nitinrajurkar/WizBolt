namespace WizBolt
{
    partial class ProjectView
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
            this.Cancel_Button = new System.Windows.Forms.Button();
            this.Ok_Button = new System.Windows.Forms.Button();
            this.FrictionCoefficient_TextBox = new System.Windows.Forms.TextBox();
            this.Detensioning_TextBox = new System.Windows.Forms.TextBox();
            this.CrossLoading_TextBox = new System.Windows.Forms.TextBox();
            this.SummaryNotes_TextBox = new System.Windows.Forms.TextBox();
            this.Notes_RichTextBox = new System.Windows.Forms.RichTextBox();
            this.Engineer_TextBox = new System.Windows.Forms.TextBox();
            this.ProjectDate_DateTimePicker = new System.Windows.Forms.DateTimePicker();
            this.Reference_TextBox = new System.Windows.Forms.TextBox();
            this.Project_TextBox = new System.Windows.Forms.TextBox();
            this.Client_TextBox = new System.Windows.Forms.TextBox();
            this.MinorDiameterArea_RadioButton = new System.Windows.Forms.RadioButton();
            this.TensileStressArea_RadioButton = new System.Windows.Forms.RadioButton();
            this.StressValue_Label = new System.Windows.Forms.Label();
            this.TorqueCoefficient_Label = new System.Windows.Forms.Label();
            this.Detensioning_Label = new System.Windows.Forms.Label();
            this.CrossLoading_Label = new System.Windows.Forms.Label();
            this.ToolRange_Label = new System.Windows.Forms.Label();
            this.SummaryDoc_Label = new System.Windows.Forms.Label();
            this.Notes_Label = new System.Windows.Forms.Label();
            this.Engineer_Label = new System.Windows.Forms.Label();
            this.StartDate_Label = new System.Windows.Forms.Label();
            this.Reference_Label = new System.Windows.Forms.Label();
            this.ProjectName_Label = new System.Windows.Forms.Label();
            this.Client_Label = new System.Windows.Forms.Label();
            this.EndDate_DateTimePicker = new System.Windows.Forms.DateTimePicker();
            this.EndDate_Label = new System.Windows.Forms.Label();
            this.Location_TextBox = new System.Windows.Forms.TextBox();
            this.Location_Label = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // Cancel_Button
            // 
            this.Cancel_Button.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Cancel_Button.ForeColor = System.Drawing.Color.DarkRed;
            this.Cancel_Button.Location = new System.Drawing.Point(456, 567);
            this.Cancel_Button.Name = "Cancel_Button";
            this.Cancel_Button.Size = new System.Drawing.Size(75, 30);
            this.Cancel_Button.TabIndex = 17;
            this.Cancel_Button.Text = "Cancel";
            this.Cancel_Button.UseVisualStyleBackColor = true;
            this.Cancel_Button.Click += new System.EventHandler(this.Cancel_Button_Click);
            // 
            // Ok_Button
            // 
            this.Ok_Button.Font = new System.Drawing.Font("Wide Latin", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Ok_Button.ForeColor = System.Drawing.Color.DarkGreen;
            this.Ok_Button.Location = new System.Drawing.Point(370, 567);
            this.Ok_Button.Name = "Ok_Button";
            this.Ok_Button.Size = new System.Drawing.Size(75, 30);
            this.Ok_Button.TabIndex = 16;
            this.Ok_Button.Text = "Ok";
            this.Ok_Button.UseVisualStyleBackColor = true;
            this.Ok_Button.Click += new System.EventHandler(this.Ok_Button_Click);
            // 
            // FrictionCoefficient_TextBox
            // 
            this.FrictionCoefficient_TextBox.Location = new System.Drawing.Point(167, 488);
            this.FrictionCoefficient_TextBox.Name = "FrictionCoefficient_TextBox";
            this.FrictionCoefficient_TextBox.Size = new System.Drawing.Size(64, 20);
            this.FrictionCoefficient_TextBox.TabIndex = 12;
            this.FrictionCoefficient_TextBox.Text = "0.12";
            this.FrictionCoefficient_TextBox.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // Detensioning_TextBox
            // 
            this.Detensioning_TextBox.Location = new System.Drawing.Point(159, 462);
            this.Detensioning_TextBox.Name = "Detensioning_TextBox";
            this.Detensioning_TextBox.Size = new System.Drawing.Size(64, 20);
            this.Detensioning_TextBox.TabIndex = 11;
            this.Detensioning_TextBox.Text = "0";
            this.Detensioning_TextBox.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // CrossLoading_TextBox
            // 
            this.CrossLoading_TextBox.Location = new System.Drawing.Point(159, 435);
            this.CrossLoading_TextBox.Name = "CrossLoading_TextBox";
            this.CrossLoading_TextBox.Size = new System.Drawing.Size(64, 20);
            this.CrossLoading_TextBox.TabIndex = 10;
            this.CrossLoading_TextBox.Text = "20";
            this.CrossLoading_TextBox.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // SummaryNotes_TextBox
            // 
            this.SummaryNotes_TextBox.Location = new System.Drawing.Point(159, 342);
            this.SummaryNotes_TextBox.MaxLength = 250;
            this.SummaryNotes_TextBox.Name = "SummaryNotes_TextBox";
            this.SummaryNotes_TextBox.Size = new System.Drawing.Size(309, 20);
            this.SummaryNotes_TextBox.TabIndex = 8;
            // 
            // Notes_RichTextBox
            // 
            this.Notes_RichTextBox.Location = new System.Drawing.Point(159, 175);
            this.Notes_RichTextBox.Name = "Notes_RichTextBox";
            this.Notes_RichTextBox.Size = new System.Drawing.Size(309, 157);
            this.Notes_RichTextBox.TabIndex = 7;
            this.Notes_RichTextBox.Text = "";
            // 
            // Engineer_TextBox
            // 
            this.Engineer_TextBox.Location = new System.Drawing.Point(159, 147);
            this.Engineer_TextBox.MaxLength = 200;
            this.Engineer_TextBox.Name = "Engineer_TextBox";
            this.Engineer_TextBox.Size = new System.Drawing.Size(309, 20);
            this.Engineer_TextBox.TabIndex = 6;
            // 
            // ProjectDate_DateTimePicker
            // 
            this.ProjectDate_DateTimePicker.Location = new System.Drawing.Point(159, 120);
            this.ProjectDate_DateTimePicker.Name = "ProjectDate_DateTimePicker";
            this.ProjectDate_DateTimePicker.Size = new System.Drawing.Size(193, 20);
            this.ProjectDate_DateTimePicker.TabIndex = 5;
            // 
            // Reference_TextBox
            // 
            this.Reference_TextBox.Location = new System.Drawing.Point(159, 94);
            this.Reference_TextBox.MaxLength = 200;
            this.Reference_TextBox.Name = "Reference_TextBox";
            this.Reference_TextBox.Size = new System.Drawing.Size(309, 20);
            this.Reference_TextBox.TabIndex = 4;
            // 
            // Project_TextBox
            // 
            this.Project_TextBox.Location = new System.Drawing.Point(159, 66);
            this.Project_TextBox.MaxLength = 200;
            this.Project_TextBox.Name = "Project_TextBox";
            this.Project_TextBox.Size = new System.Drawing.Size(309, 20);
            this.Project_TextBox.TabIndex = 3;
            // 
            // Client_TextBox
            // 
            this.Client_TextBox.Location = new System.Drawing.Point(159, 16);
            this.Client_TextBox.MaxLength = 200;
            this.Client_TextBox.Name = "Client_TextBox";
            this.Client_TextBox.Size = new System.Drawing.Size(309, 20);
            this.Client_TextBox.TabIndex = 1;
            // 
            // MinorDiameterArea_RadioButton
            // 
            this.MinorDiameterArea_RadioButton.AutoSize = true;
            this.MinorDiameterArea_RadioButton.Location = new System.Drawing.Point(306, 513);
            this.MinorDiameterArea_RadioButton.Name = "MinorDiameterArea_RadioButton";
            this.MinorDiameterArea_RadioButton.Size = new System.Drawing.Size(121, 17);
            this.MinorDiameterArea_RadioButton.TabIndex = 14;
            this.MinorDiameterArea_RadioButton.Text = "Minor Diameter Area";
            this.MinorDiameterArea_RadioButton.UseVisualStyleBackColor = true;
            // 
            // TensileStressArea_RadioButton
            // 
            this.TensileStressArea_RadioButton.AutoSize = true;
            this.TensileStressArea_RadioButton.Checked = true;
            this.TensileStressArea_RadioButton.Location = new System.Drawing.Point(167, 513);
            this.TensileStressArea_RadioButton.Name = "TensileStressArea_RadioButton";
            this.TensileStressArea_RadioButton.Size = new System.Drawing.Size(116, 17);
            this.TensileStressArea_RadioButton.TabIndex = 13;
            this.TensileStressArea_RadioButton.TabStop = true;
            this.TensileStressArea_RadioButton.Text = "Tensile Stress Area";
            this.TensileStressArea_RadioButton.UseVisualStyleBackColor = true;
            // 
            // StressValue_Label
            // 
            this.StressValue_Label.AutoSize = true;
            this.StressValue_Label.Location = new System.Drawing.Point(18, 515);
            this.StressValue_Label.Name = "StressValue_Label";
            this.StressValue_Label.Size = new System.Drawing.Size(138, 13);
            this.StressValue_Label.TabIndex = 39;
            this.StressValue_Label.Text = "Stress values are based on:";
            // 
            // TorqueCoefficient_Label
            // 
            this.TorqueCoefficient_Label.AutoSize = true;
            this.TorqueCoefficient_Label.Location = new System.Drawing.Point(13, 491);
            this.TorqueCoefficient_Label.Name = "TorqueCoefficient_Label";
            this.TorqueCoefficient_Label.Size = new System.Drawing.Size(155, 13);
            this.TorqueCoefficient_Label.TabIndex = 38;
            this.TorqueCoefficient_Label.Text = "Torque Coefficient of Friction µ:";
            // 
            // Detensioning_Label
            // 
            this.Detensioning_Label.AutoSize = true;
            this.Detensioning_Label.Location = new System.Drawing.Point(76, 465);
            this.Detensioning_Label.Name = "Detensioning_Label";
            this.Detensioning_Label.Size = new System.Drawing.Size(83, 13);
            this.Detensioning_Label.TabIndex = 37;
            this.Detensioning_Label.Text = "Detensioning %:";
            // 
            // CrossLoading_Label
            // 
            this.CrossLoading_Label.AutoSize = true;
            this.CrossLoading_Label.Location = new System.Drawing.Point(71, 438);
            this.CrossLoading_Label.Name = "CrossLoading_Label";
            this.CrossLoading_Label.Size = new System.Drawing.Size(88, 13);
            this.CrossLoading_Label.TabIndex = 36;
            this.CrossLoading_Label.Text = "Cross Loading %:";
            // 
            // ToolRange_Label
            // 
            this.ToolRange_Label.AutoSize = true;
            this.ToolRange_Label.Location = new System.Drawing.Point(78, 378);
            this.ToolRange_Label.Name = "ToolRange_Label";
            this.ToolRange_Label.Size = new System.Drawing.Size(66, 13);
            this.ToolRange_Label.TabIndex = 35;
            this.ToolRange_Label.Text = "Tool Range:";
            // 
            // SummaryDoc_Label
            // 
            this.SummaryDoc_Label.AutoSize = true;
            this.SummaryDoc_Label.Location = new System.Drawing.Point(49, 345);
            this.SummaryDoc_Label.Name = "SummaryDoc_Label";
            this.SummaryDoc_Label.Size = new System.Drawing.Size(110, 13);
            this.SummaryDoc_Label.TabIndex = 34;
            this.SummaryDoc_Label.Text = "Summary Doc. Notes:";
            // 
            // Notes_Label
            // 
            this.Notes_Label.AutoSize = true;
            this.Notes_Label.Location = new System.Drawing.Point(121, 172);
            this.Notes_Label.Name = "Notes_Label";
            this.Notes_Label.Size = new System.Drawing.Size(38, 13);
            this.Notes_Label.TabIndex = 33;
            this.Notes_Label.Text = "Notes:";
            // 
            // Engineer_Label
            // 
            this.Engineer_Label.AutoSize = true;
            this.Engineer_Label.Location = new System.Drawing.Point(107, 150);
            this.Engineer_Label.Name = "Engineer_Label";
            this.Engineer_Label.Size = new System.Drawing.Size(52, 13);
            this.Engineer_Label.TabIndex = 32;
            this.Engineer_Label.Text = "Engineer:";
            // 
            // StartDate_Label
            // 
            this.StartDate_Label.AutoSize = true;
            this.StartDate_Label.Location = new System.Drawing.Point(99, 124);
            this.StartDate_Label.Name = "StartDate_Label";
            this.StartDate_Label.Size = new System.Drawing.Size(58, 13);
            this.StartDate_Label.TabIndex = 31;
            this.StartDate_Label.Text = "Start Date:";
            // 
            // Reference_Label
            // 
            this.Reference_Label.AutoSize = true;
            this.Reference_Label.Location = new System.Drawing.Point(99, 97);
            this.Reference_Label.Name = "Reference_Label";
            this.Reference_Label.Size = new System.Drawing.Size(60, 13);
            this.Reference_Label.TabIndex = 30;
            this.Reference_Label.Text = "Reference:";
            // 
            // ProjectName_Label
            // 
            this.ProjectName_Label.AutoSize = true;
            this.ProjectName_Label.Location = new System.Drawing.Point(116, 69);
            this.ProjectName_Label.Name = "ProjectName_Label";
            this.ProjectName_Label.Size = new System.Drawing.Size(43, 13);
            this.ProjectName_Label.TabIndex = 29;
            this.ProjectName_Label.Text = "Project:";
            // 
            // Client_Label
            // 
            this.Client_Label.AutoSize = true;
            this.Client_Label.Location = new System.Drawing.Point(123, 19);
            this.Client_Label.Name = "Client_Label";
            this.Client_Label.Size = new System.Drawing.Size(36, 13);
            this.Client_Label.TabIndex = 28;
            this.Client_Label.Text = "Client:";
            // 
            // EndDate_DateTimePicker
            // 
            this.EndDate_DateTimePicker.Location = new System.Drawing.Point(134, 539);
            this.EndDate_DateTimePicker.Name = "EndDate_DateTimePicker";
            this.EndDate_DateTimePicker.Size = new System.Drawing.Size(205, 20);
            this.EndDate_DateTimePicker.TabIndex = 15;
            // 
            // EndDate_Label
            // 
            this.EndDate_Label.AutoSize = true;
            this.EndDate_Label.Location = new System.Drawing.Point(71, 543);
            this.EndDate_Label.Name = "EndDate_Label";
            this.EndDate_Label.Size = new System.Drawing.Size(55, 13);
            this.EndDate_Label.TabIndex = 55;
            this.EndDate_Label.Text = "End Date:";
            // 
            // Location_TextBox
            // 
            this.Location_TextBox.Location = new System.Drawing.Point(159, 40);
            this.Location_TextBox.MaxLength = 200;
            this.Location_TextBox.Name = "Location_TextBox";
            this.Location_TextBox.Size = new System.Drawing.Size(309, 20);
            this.Location_TextBox.TabIndex = 2;
            // 
            // Location_Label
            // 
            this.Location_Label.AutoSize = true;
            this.Location_Label.Location = new System.Drawing.Point(109, 43);
            this.Location_Label.Name = "Location_Label";
            this.Location_Label.Size = new System.Drawing.Size(51, 13);
            this.Location_Label.TabIndex = 57;
            this.Location_Label.Text = "Location:";
            // 
            // ProjectView
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(544, 602);
            this.Controls.Add(this.Location_TextBox);
            this.Controls.Add(this.Location_Label);
            this.Controls.Add(this.EndDate_DateTimePicker);
            this.Controls.Add(this.EndDate_Label);
            this.Controls.Add(this.Cancel_Button);
            this.Controls.Add(this.Ok_Button);
            this.Controls.Add(this.FrictionCoefficient_TextBox);
            this.Controls.Add(this.Detensioning_TextBox);
            this.Controls.Add(this.CrossLoading_TextBox);
            this.Controls.Add(this.SummaryNotes_TextBox);
            this.Controls.Add(this.Notes_RichTextBox);
            this.Controls.Add(this.Engineer_TextBox);
            this.Controls.Add(this.ProjectDate_DateTimePicker);
            this.Controls.Add(this.Reference_TextBox);
            this.Controls.Add(this.Project_TextBox);
            this.Controls.Add(this.Client_TextBox);
            this.Controls.Add(this.MinorDiameterArea_RadioButton);
            this.Controls.Add(this.TensileStressArea_RadioButton);
            this.Controls.Add(this.StressValue_Label);
            this.Controls.Add(this.TorqueCoefficient_Label);
            this.Controls.Add(this.Detensioning_Label);
            this.Controls.Add(this.CrossLoading_Label);
            this.Controls.Add(this.ToolRange_Label);
            this.Controls.Add(this.SummaryDoc_Label);
            this.Controls.Add(this.Notes_Label);
            this.Controls.Add(this.Engineer_Label);
            this.Controls.Add(this.StartDate_Label);
            this.Controls.Add(this.Reference_Label);
            this.Controls.Add(this.ProjectName_Label);
            this.Controls.Add(this.Client_Label);
            this.Name = "ProjectView";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Project Details";
            this.Load += new System.EventHandler(this.ProjectView_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button Cancel_Button;
        private System.Windows.Forms.Button Ok_Button;
        private System.Windows.Forms.TextBox FrictionCoefficient_TextBox;
        private System.Windows.Forms.TextBox Detensioning_TextBox;
        private System.Windows.Forms.TextBox CrossLoading_TextBox;
        private System.Windows.Forms.TextBox SummaryNotes_TextBox;
        private System.Windows.Forms.RichTextBox Notes_RichTextBox;
        private System.Windows.Forms.TextBox Engineer_TextBox;
        private System.Windows.Forms.DateTimePicker ProjectDate_DateTimePicker;
        private System.Windows.Forms.TextBox Reference_TextBox;
        private System.Windows.Forms.TextBox Project_TextBox;
        private System.Windows.Forms.TextBox Client_TextBox;
        private System.Windows.Forms.RadioButton MinorDiameterArea_RadioButton;
        private System.Windows.Forms.RadioButton TensileStressArea_RadioButton;
        private System.Windows.Forms.Label StressValue_Label;
        private System.Windows.Forms.Label TorqueCoefficient_Label;
        private System.Windows.Forms.Label Detensioning_Label;
        private System.Windows.Forms.Label CrossLoading_Label;
        private System.Windows.Forms.Label ToolRange_Label;
        private System.Windows.Forms.Label SummaryDoc_Label;
        private System.Windows.Forms.Label Notes_Label;
        private System.Windows.Forms.Label Engineer_Label;
        private System.Windows.Forms.Label StartDate_Label;
        private System.Windows.Forms.Label Reference_Label;
        private System.Windows.Forms.Label ProjectName_Label;
        private System.Windows.Forms.Label Client_Label;
        private System.Windows.Forms.DateTimePicker EndDate_DateTimePicker;
        private System.Windows.Forms.Label EndDate_Label;
        private System.Windows.Forms.TextBox Location_TextBox;
        private System.Windows.Forms.Label Location_Label;
    }
}