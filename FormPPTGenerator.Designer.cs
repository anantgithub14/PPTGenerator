namespace PPTGenerator
{
    partial class FormPPTGenerator
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
            this.openWordFileDialog = new System.Windows.Forms.OpenFileDialog();
            this.btnBrowseWordFile = new System.Windows.Forms.Button();
            this.txtWordFilePath = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.btnPPTGenerator = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.txtPPTFilePath = new System.Windows.Forms.TextBox();
            this.btnBrowsePPTFile = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // openWordFileDialog
            // 
            this.openWordFileDialog.FileName = "openWordFileDialog";
            // 
            // btnBrowseWordFile
            // 
            this.btnBrowseWordFile.Location = new System.Drawing.Point(434, 20);
            this.btnBrowseWordFile.Name = "btnBrowseWordFile";
            this.btnBrowseWordFile.Size = new System.Drawing.Size(25, 23);
            this.btnBrowseWordFile.TabIndex = 0;
            this.btnBrowseWordFile.Text = "...";
            this.btnBrowseWordFile.UseVisualStyleBackColor = true;
            this.btnBrowseWordFile.Click += new System.EventHandler(this.btnBrowseWordFile_Click);
            // 
            // txtWordFilePath
            // 
            this.txtWordFilePath.Location = new System.Drawing.Point(204, 61);
            this.txtWordFilePath.Name = "txtWordFilePath";
            this.txtWordFilePath.Size = new System.Drawing.Size(243, 20);
            this.txtWordFilePath.TabIndex = 1;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(111, 64);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(93, 13);
            this.label1.TabIndex = 2;
            this.label1.Text = "Browse Word File:";
            // 
            // btnPPTGenerator
            // 
            this.btnPPTGenerator.Location = new System.Drawing.Point(250, 155);
            this.btnPPTGenerator.Name = "btnPPTGenerator";
            this.btnPPTGenerator.Size = new System.Drawing.Size(97, 23);
            this.btnPPTGenerator.TabIndex = 3;
            this.btnPPTGenerator.Text = "Generate PPT";
            this.btnPPTGenerator.UseVisualStyleBackColor = true;
            this.btnPPTGenerator.Click += new System.EventHandler(this.btnPPTGenerator_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(55, 104);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(149, 13);
            this.label2.TabIndex = 6;
            this.label2.Text = "Browse PowerPoint Template:";
            // 
            // txtPPTFilePath
            // 
            this.txtPPTFilePath.Location = new System.Drawing.Point(204, 101);
            this.txtPPTFilePath.Name = "txtPPTFilePath";
            this.txtPPTFilePath.Size = new System.Drawing.Size(243, 20);
            this.txtPPTFilePath.TabIndex = 5;
            // 
            // btnBrowsePPTFile
            // 
            this.btnBrowsePPTFile.Location = new System.Drawing.Point(434, 61);
            this.btnBrowsePPTFile.Name = "btnBrowsePPTFile";
            this.btnBrowsePPTFile.Size = new System.Drawing.Size(25, 23);
            this.btnBrowsePPTFile.TabIndex = 4;
            this.btnBrowsePPTFile.Text = "...";
            this.btnBrowsePPTFile.UseVisualStyleBackColor = true;
            this.btnBrowsePPTFile.Click += new System.EventHandler(this.btnBrowsePPTFile_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.btnBrowsePPTFile);
            this.groupBox1.Controls.Add(this.btnBrowseWordFile);
            this.groupBox1.Location = new System.Drawing.Point(17, 38);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(524, 112);
            this.groupBox1.TabIndex = 7;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Input Templates";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(599, 213);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.txtPPTFilePath);
            this.Controls.Add(this.btnPPTGenerator);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.txtWordFilePath);
            this.Controls.Add(this.groupBox1);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "Form1";
            this.Text = "PPTGenerator";
            this.groupBox1.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.OpenFileDialog openWordFileDialog;
        private System.Windows.Forms.Button btnBrowseWordFile;
        private System.Windows.Forms.TextBox txtWordFilePath;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btnPPTGenerator;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtPPTFilePath;
        private System.Windows.Forms.Button btnBrowsePPTFile;
        private System.Windows.Forms.GroupBox groupBox1;
    }
}

