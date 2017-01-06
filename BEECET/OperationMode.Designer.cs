//
// (C) Copyright 2003-2010 by Autodesk, Inc.
//
// Permission to use, copy, modify, and distribute this software in
// object code form for any purpose and without fee is hereby granted,
// provided that the above copyright notice appears in all copies and
// that both that copyright notice and the limited warranty and
// restricted rights notice below appear in all supporting
// documentation.
//
// AUTODESK PROVIDES THIS PROGRAM "AS IS" AND WITH ALL FAULTS.
// AUTODESK SPECIFICALLY DISCLAIMS ANY IMPLIED WARRANTY OF
// MERCHANTABILITY OR FITNESS FOR A PARTICULAR USE. AUTODESK, INC.
// DOES NOT WARRANT THAT THE OPERATION OF THE PROGRAM WILL BE
// UNINTERRUPTED OR ERROR FREE.
//
// Use, duplication, or disclosure by the U.S. Government is subject to
// restrictions set forth in FAR 52.227-19 (Commercial Computer
// Software - Restricted Rights) and DFAR 252.227-7013(c)(1)(ii)
// (Rights in Technical Data and Computer Software), as applicable.
//


namespace Revit.SDK.Samples.AnalyticalSupportData_Info.CS
{
    partial class OperationMode
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
            this.closeButton = new System.Windows.Forms.Button();
            this.ProceedButton = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.EmbodiedECAnalyisisRadioButton = new System.Windows.Forms.RadioButton();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.ProjectLocationTextBox = new System.Windows.Forms.TextBox();
            this.DesignOptionNoTextBox = new System.Windows.Forms.TextBox();
            this.DesignLifeTextBox = new System.Windows.Forms.TextBox();
            this.ProjectTitleTextBox = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.ProjectIDTextBox = new System.Windows.Forms.TextBox();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // closeButton
            // 
            this.closeButton.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.closeButton.Location = new System.Drawing.Point(278, 486);
            this.closeButton.Margin = new System.Windows.Forms.Padding(2);
            this.closeButton.Name = "closeButton";
            this.closeButton.Size = new System.Drawing.Size(68, 24);
            this.closeButton.TabIndex = 5;
            this.closeButton.Text = "&Close";
            this.closeButton.UseVisualStyleBackColor = true;
            this.closeButton.Click += new System.EventHandler(this.closeButton_Click);
            // 
            // ProceedButton
            // 
            this.ProceedButton.Location = new System.Drawing.Point(168, 200);
            this.ProceedButton.Name = "ProceedButton";
            this.ProceedButton.Size = new System.Drawing.Size(73, 24);
            this.ProceedButton.TabIndex = 2;
            this.ProceedButton.Text = "Proceed";
            this.ProceedButton.UseVisualStyleBackColor = true;
            this.ProceedButton.Click += new System.EventHandler(this.ProceedButton_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.EmbodiedECAnalyisisRadioButton);
            this.groupBox1.Controls.Add(this.ProceedButton);
            this.groupBox1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.groupBox1.Location = new System.Drawing.Point(35, 222);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(311, 238);
            this.groupBox1.TabIndex = 4;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Operation mode ";
            // 
            // EmbodiedECAnalyisisRadioButton
            // 
            this.EmbodiedECAnalyisisRadioButton.AccessibleRole = System.Windows.Forms.AccessibleRole.Pane;
            this.EmbodiedECAnalyisisRadioButton.AutoSize = true;
            this.EmbodiedECAnalyisisRadioButton.Location = new System.Drawing.Point(11, 84);
            this.EmbodiedECAnalyisisRadioButton.Name = "EmbodiedECAnalyisisRadioButton";
            this.EmbodiedECAnalyisisRadioButton.Size = new System.Drawing.Size(230, 19);
            this.EmbodiedECAnalyisisRadioButton.TabIndex = 5;
            this.EmbodiedECAnalyisisRadioButton.TabStop = true;
            this.EmbodiedECAnalyisisRadioButton.Text = "Emodied Energy and Carbon Analysis";
            this.EmbodiedECAnalyisisRadioButton.UseVisualStyleBackColor = true;
            this.EmbodiedECAnalyisisRadioButton.CheckedChanged += new System.EventHandler(this.EmbodiedECAnalyisisRadioButton_CheckedChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(37, 44);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(54, 13);
            this.label1.TabIndex = 6;
            this.label1.Text = "Project ID";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(37, 144);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(114, 13);
            this.label2.TabIndex = 9;
            this.label2.Text = "Design Option Number";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(37, 105);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(80, 13);
            this.label3.TabIndex = 8;
            this.label3.Text = "Project location";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(40, 174);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(96, 13);
            this.label4.TabIndex = 10;
            this.label4.Text = "Design Life (Years)";
            // 
            // ProjectLocationTextBox
            // 
            this.ProjectLocationTextBox.Location = new System.Drawing.Point(170, 102);
            this.ProjectLocationTextBox.Name = "ProjectLocationTextBox";
            this.ProjectLocationTextBox.Size = new System.Drawing.Size(176, 20);
            this.ProjectLocationTextBox.TabIndex = 2;
            this.ProjectLocationTextBox.Text = "Nottingham, UK";
            // 
            // DesignOptionNoTextBox
            // 
            this.DesignOptionNoTextBox.DataBindings.Add(new System.Windows.Forms.Binding("Text", global::AnalyticalSupportData_Info.Properties.Settings.Default, "DON", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            this.DesignOptionNoTextBox.Location = new System.Drawing.Point(170, 137);
            this.DesignOptionNoTextBox.Name = "DesignOptionNoTextBox";
            this.DesignOptionNoTextBox.Size = new System.Drawing.Size(176, 20);
            this.DesignOptionNoTextBox.TabIndex = 3;
            this.DesignOptionNoTextBox.Text = global::AnalyticalSupportData_Info.Properties.Settings.Default.DON;
            // 
            // DesignLifeTextBox
            // 
            this.DesignLifeTextBox.Location = new System.Drawing.Point(170, 171);
            this.DesignLifeTextBox.Name = "DesignLifeTextBox";
            this.DesignLifeTextBox.Size = new System.Drawing.Size(176, 20);
            this.DesignLifeTextBox.TabIndex = 4;
            this.DesignLifeTextBox.Text = "80";
            // 
            // ProjectTitleTextBox
            // 
            this.ProjectTitleTextBox.Location = new System.Drawing.Point(169, 69);
            this.ProjectTitleTextBox.Name = "ProjectTitleTextBox";
            this.ProjectTitleTextBox.Size = new System.Drawing.Size(176, 20);
            this.ProjectTitleTextBox.TabIndex = 1;
            this.ProjectTitleTextBox.Text = "Research Building Project";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(36, 76);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(63, 13);
            this.label5.TabIndex = 7;
            this.label5.Text = "Project Title";
            // 
            // ProjectIDTextBox
            // 
            this.ProjectIDTextBox.DataBindings.Add(new System.Windows.Forms.Binding("Text", global::AnalyticalSupportData_Info.Properties.Settings.Default, "PID", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            this.ProjectIDTextBox.Location = new System.Drawing.Point(170, 37);
            this.ProjectIDTextBox.Name = "ProjectIDTextBox";
            this.ProjectIDTextBox.Size = new System.Drawing.Size(176, 20);
            this.ProjectIDTextBox.TabIndex = 0;
            this.ProjectIDTextBox.Text = global::AnalyticalSupportData_Info.Properties.Settings.Default.PID;
            // 
            // OperationMode
            // 
            this.AcceptButton = this.closeButton;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.closeButton;
            this.ClientSize = new System.Drawing.Size(377, 522);
            this.Controls.Add(this.ProjectTitleTextBox);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.DesignLifeTextBox);
            this.Controls.Add(this.DesignOptionNoTextBox);
            this.Controls.Add(this.ProjectLocationTextBox);
            this.Controls.Add(this.ProjectIDTextBox);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.closeButton);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Margin = new System.Windows.Forms.Padding(2);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "OperationMode";
            this.ShowInTaskbar = false;
            this.Text = "STEEL SUSTAINABILITY ESTIMATOR";
            this.Load += new System.EventHandler(this.OperationMode_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button closeButton;
        private System.Windows.Forms.Button ProceedButton;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox ProjectIDTextBox;
        private System.Windows.Forms.TextBox ProjectLocationTextBox;
        private System.Windows.Forms.TextBox DesignOptionNoTextBox;
        private System.Windows.Forms.TextBox DesignLifeTextBox;
        private System.Windows.Forms.TextBox ProjectTitleTextBox;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.RadioButton EmbodiedECAnalyisisRadioButton;
    }
}