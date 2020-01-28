namespace PPTGenerator
{
    partial class Form1
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
            this.SuspendLayout();
            // 
            // openWordFileDialog
            // 
            this.openWordFileDialog.FileName = "openWordFileDialog";
            // 
            // btnBrowseWordFile
            // 
            this.btnBrowseWordFile.Location = new System.Drawing.Point(410, 61);
            this.btnBrowseWordFile.Name = "btnBrowseWordFile";
            this.btnBrowseWordFile.Size = new System.Drawing.Size(75, 23);
            this.btnBrowseWordFile.TabIndex = 0;
            this.btnBrowseWordFile.Text = "Browse";
            this.btnBrowseWordFile.UseVisualStyleBackColor = true;
            this.btnBrowseWordFile.Click += new System.EventHandler(this.btnBrowseWordFile_Click);
            // 
            // txtWordFilePath
            // 
            this.txtWordFilePath.Location = new System.Drawing.Point(161, 61);
            this.txtWordFilePath.Name = "txtWordFilePath";
            this.txtWordFilePath.Size = new System.Drawing.Size(243, 20);
            this.txtWordFilePath.TabIndex = 1;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(70, 64);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(85, 13);
            this.label1.TabIndex = 2;
            this.label1.Text = "Select Word File";
            // 
            // btnPPTGenerator
            // 
            this.btnPPTGenerator.Location = new System.Drawing.Point(250, 110);
            this.btnPPTGenerator.Name = "btnPPTGenerator";
            this.btnPPTGenerator.Size = new System.Drawing.Size(97, 23);
            this.btnPPTGenerator.TabIndex = 3;
            this.btnPPTGenerator.Text = "Generate PPT";
            this.btnPPTGenerator.UseVisualStyleBackColor = true;
            this.btnPPTGenerator.Click += new System.EventHandler(this.btnPPTGenerator_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.btnPPTGenerator);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.txtWordFilePath);
            this.Controls.Add(this.btnBrowseWordFile);
            this.Name = "Form1";
            this.Text = "Form1";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.OpenFileDialog openWordFileDialog;
        private System.Windows.Forms.Button btnBrowseWordFile;
        private System.Windows.Forms.TextBox txtWordFilePath;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btnPPTGenerator;
    }
}

