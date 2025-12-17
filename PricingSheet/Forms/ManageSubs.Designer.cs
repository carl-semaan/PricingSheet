using DocumentFormat.OpenXml.Office2010.ExcelAc;
using PricingSheet.Models;
using System.Collections.Generic;

namespace PricingSheet.Forms
{
    partial class ManageSubs
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
            this.checkedListBox1 = new System.Windows.Forms.CheckedListBox();
            this.eventLog1 = new System.Diagnostics.EventLog();
            this.eventLog2 = new System.Diagnostics.EventLog();
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.button1 = new System.Windows.Forms.Button();
            this.button3 = new System.Windows.Forms.Button();
            this.Instr = new System.Windows.Forms.Label();
            this.Mat = new System.Windows.Forms.Label();
            this.Subscriptions = new System.Windows.Forms.Label();
            this.button2 = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.eventLog1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.eventLog2)).BeginInit();
            this.SuspendLayout();
            // 
            // checkedListBox1
            // 
            this.checkedListBox1.BackColor = System.Drawing.SystemColors.Window;
            this.checkedListBox1.CheckOnClick = true;
            this.checkedListBox1.FormattingEnabled = true;
            this.checkedListBox1.Location = new System.Drawing.Point(12, 12);
            this.checkedListBox1.Name = "checkedListBox1";
            this.checkedListBox1.Size = new System.Drawing.Size(121, 154);
            this.checkedListBox1.TabIndex = 0;
            this.checkedListBox1.ItemCheck += new System.Windows.Forms.ItemCheckEventHandler(this.CheckedListBox1_ItemCheck);
            // 
            // eventLog1
            // 
            this.eventLog1.SynchronizingObject = this;
            // 
            // eventLog2
            // 
            this.eventLog2.SynchronizingObject = this;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(142, 12);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 23);
            this.button1.TabIndex = 1;
            this.button1.Text = "Select All";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(223, 143);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(75, 23);
            this.button3.TabIndex = 3;
            this.button3.Text = "Done";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // Instr
            // 
            this.Instr.AutoSize = true;
            this.Instr.Location = new System.Drawing.Point(139, 47);
            this.Instr.Name = "Instr";
            this.Instr.Size = new System.Drawing.Size(73, 13);
            this.Instr.TabIndex = 4;
            this.Instr.Text = "Instruments: 0";
            // 
            // Mat
            // 
            this.Mat.AutoSize = true;
            this.Mat.Location = new System.Drawing.Point(139, 69);
            this.Mat.Name = "Mat";
            this.Mat.Size = new System.Drawing.Size(64, 13);
            this.Mat.TabIndex = 5;
            this.Mat.Text = "Maturities: 0";
            // 
            // Subscriptions
            // 
            this.Subscriptions.AutoSize = true;
            this.Subscriptions.Location = new System.Drawing.Point(139, 91);
            this.Subscriptions.Name = "Subscriptions";
            this.Subscriptions.Size = new System.Drawing.Size(93, 13);
            this.Subscriptions.TabIndex = 6;
            this.Subscriptions.Text = "Subscriptions: 0/0";
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(223, 12);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(75, 23);
            this.button2.TabIndex = 7;
            this.button2.Text = "Clear All";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // ManageSubs
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(307, 183);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.Subscriptions);
            this.Controls.Add(this.Mat);
            this.Controls.Add(this.Instr);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.checkedListBox1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.Name = "ManageSubs";
            this.Text = "Manage Subscriptions";
            ((System.ComponentModel.ISupportInitialize)(this.eventLog1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.eventLog2)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }
        #endregion

        private System.Windows.Forms.CheckedListBox checkedListBox1;
        private System.Diagnostics.EventLog eventLog1;
        private System.Diagnostics.EventLog eventLog2;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
        private System.Windows.Forms.Label Instr;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Label Mat;
        private System.Windows.Forms.Label Subscriptions;
        private System.Windows.Forms.Button button2;
    }
}