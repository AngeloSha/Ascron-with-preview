namespace EmailToPdfConverter
{
    partial class Form1
    {
        private System.ComponentModel.IContainer components = null;
        private System.Windows.Forms.Button btnQueue1;
        private System.Windows.Forms.Button btnQueue2;
        private System.Windows.Forms.Button btnQueue3;
        private System.Windows.Forms.Button btnQueue4;
        private System.Windows.Forms.Button btnQueue5;
        private System.Windows.Forms.Button btnQueue6;
        private System.Windows.Forms.Button btnMakePreview;
        private System.Windows.Forms.CheckBox chkTopMost;

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
            this.btnQueue1 = new System.Windows.Forms.Button();
            this.btnQueue2 = new System.Windows.Forms.Button();
            this.btnQueue3 = new System.Windows.Forms.Button();
            this.btnQueue4 = new System.Windows.Forms.Button();
            this.btnQueue5 = new System.Windows.Forms.Button();
            this.btnQueue6 = new System.Windows.Forms.Button();
            this.btnMakePreview = new System.Windows.Forms.Button();
            this.chkTopMost = new System.Windows.Forms.CheckBox();
            this.SuspendLayout();
            // 
            // btnQueue1
            // 
            this.btnQueue1.Location = new System.Drawing.Point(12, 41);
            this.btnQueue1.Name = "btnQueue1";
            this.btnQueue1.Size = new System.Drawing.Size(75, 23);
            this.btnQueue1.TabIndex = 0;
            this.btnQueue1.Text = "Queue 1";
            this.btnQueue1.UseVisualStyleBackColor = true;
            this.btnQueue1.Click += new System.EventHandler(this.btnQueue1_Click);
            // 
            // btnQueue2
            // 
            this.btnQueue2.Location = new System.Drawing.Point(93, 41);
            this.btnQueue2.Name = "btnQueue2";
            this.btnQueue2.Size = new System.Drawing.Size(75, 23);
            this.btnQueue2.TabIndex = 1;
            this.btnQueue2.Text = "Queue 2";
            this.btnQueue2.UseVisualStyleBackColor = true;
            this.btnQueue2.Click += new System.EventHandler(this.btnQueue2_Click);
            // 
            // btnQueue3
            // 
            this.btnQueue3.Location = new System.Drawing.Point(174, 41);
            this.btnQueue3.Name = "btnQueue3";
            this.btnQueue3.Size = new System.Drawing.Size(75, 23);
            this.btnQueue3.TabIndex = 2;
            this.btnQueue3.Text = "Queue 3";
            this.btnQueue3.UseVisualStyleBackColor = true;
            this.btnQueue3.Click += new System.EventHandler(this.btnQueue3_Click);
            // 
            // btnQueue4
            // 
            this.btnQueue4.Location = new System.Drawing.Point(255, 41);
            this.btnQueue4.Name = "btnQueue4";
            this.btnQueue4.Size = new System.Drawing.Size(75, 23);
            this.btnQueue4.TabIndex = 3;
            this.btnQueue4.Text = "Queue 4";
            this.btnQueue4.UseVisualStyleBackColor = true;
            this.btnQueue4.Click += new System.EventHandler(this.btnQueue4_Click);
            // 
            // btnQueue5
            // 
            this.btnQueue5.Location = new System.Drawing.Point(336, 41);
            this.btnQueue5.Name = "btnQueue5";
            this.btnQueue5.Size = new System.Drawing.Size(75, 23);
            this.btnQueue5.TabIndex = 4;
            this.btnQueue5.Text = "Queue 5";
            this.btnQueue5.UseVisualStyleBackColor = true;
            this.btnQueue5.Click += new System.EventHandler(this.btnQueue5_Click);
            // 
            // btnQueue6
            // 
            this.btnQueue6.Location = new System.Drawing.Point(417, 41);
            this.btnQueue6.Name = "btnQueue6";
            this.btnQueue6.Size = new System.Drawing.Size(75, 23);
            this.btnQueue6.TabIndex = 5;
            this.btnQueue6.Text = "Queue 6";
            this.btnQueue6.UseVisualStyleBackColor = true;
            this.btnQueue6.Click += new System.EventHandler(this.btnQueue6_Click);
            // 
            // btnMakePreview
            // 
            this.btnMakePreview.Location = new System.Drawing.Point(12, 12);
            this.btnMakePreview.Name = "btnMakePreview";
            this.btnMakePreview.Size = new System.Drawing.Size(100, 23);
            this.btnMakePreview.TabIndex = 7;
            this.btnMakePreview.Text = "Make Preview";
            this.btnMakePreview.UseVisualStyleBackColor = true;
            this.btnMakePreview.Click += new System.EventHandler(this.btnMakePreview_Click);
            // 
            // chkTopMost
            // 
            this.chkTopMost.AutoSize = true;
            this.chkTopMost.Location = new System.Drawing.Point(12, 70);
            this.chkTopMost.Name = "chkTopMost";
            this.chkTopMost.Size = new System.Drawing.Size(94, 19);
            this.chkTopMost.TabIndex = 6;
            this.chkTopMost.Text = "Keep on Top";
            this.chkTopMost.UseVisualStyleBackColor = true;
            this.chkTopMost.CheckedChanged += new System.EventHandler(this.chkTopMost_CheckedChanged);
            // 
            // Form1
            // 
            this.ClientSize = new System.Drawing.Size(800, 600);
            this.Controls.Add(this.chkTopMost);
            this.Controls.Add(this.btnMakePreview);
            this.Controls.Add(this.btnQueue6);
            this.Controls.Add(this.btnQueue5);
            this.Controls.Add(this.btnQueue4);
            this.Controls.Add(this.btnQueue3);
            this.Controls.Add(this.btnQueue2);
            this.Controls.Add(this.btnQueue1);
            this.Name = "Form1";
            this.Text = "Email to PDF Converter";
            this.ResumeLayout(false);
            this.PerformLayout();
        }

        #endregion
    }
}
