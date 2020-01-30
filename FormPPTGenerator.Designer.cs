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
            this.btnBrowsePPTFile = new System.Windows.Forms.Button();
            this.btnBrowseWordFile = new System.Windows.Forms.Button();
            this.btnPPTGenerator = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.txtPPTFilePath = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.txtWordFilePath = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // openWordFileDialog
            // 
            this.openWordFileDialog.FileName = "openWordFileDialog";
            // 
            // btnBrowsePPTFile
            // 
            this.btnBrowsePPTFile.Location = new System.Drawing.Point(430, 142);
            this.btnBrowsePPTFile.Name = "btnBrowsePPTFile";
            this.btnBrowsePPTFile.Size = new System.Drawing.Size(25, 23);
            this.btnBrowsePPTFile.TabIndex = 11;
            this.btnBrowsePPTFile.Text = "...";
            this.btnBrowsePPTFile.UseVisualStyleBackColor = true;
            this.btnBrowsePPTFile.Click += new System.EventHandler(this.btnBrowsePPTFile_Click);
            // 
            // btnBrowseWordFile
            // 
            this.btnBrowseWordFile.Location = new System.Drawing.Point(430, 105);
            this.btnBrowseWordFile.Name = "btnBrowseWordFile";
            this.btnBrowseWordFile.Size = new System.Drawing.Size(25, 23);
            this.btnBrowseWordFile.TabIndex = 7;
            this.btnBrowseWordFile.Text = "...";
            this.btnBrowseWordFile.UseVisualStyleBackColor = true;
            this.btnBrowseWordFile.Click += new System.EventHandler(this.btnBrowseWordFile_Click);
            // 
            // btnPPTGenerator
            // 
            this.btnPPTGenerator.Location = new System.Drawing.Point(183, 179);
            this.btnPPTGenerator.Name = "btnPPTGenerator";
            this.btnPPTGenerator.Size = new System.Drawing.Size(97, 23);
            this.btnPPTGenerator.TabIndex = 10;
            this.btnPPTGenerator.Text = "Generate PPT";
            this.btnPPTGenerator.UseVisualStyleBackColor = true;
            this.btnPPTGenerator.Click += new System.EventHandler(this.btnPPTGenerator_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.BackColor = System.Drawing.Color.Transparent;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.ForeColor = System.Drawing.SystemColors.Highlight;
            this.label2.Location = new System.Drawing.Point(1, 147);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(176, 13);
            this.label2.TabIndex = 13;
            this.label2.Text = "Browse PowerPoint Template:";
            // 
            // txtPPTFilePath
            // 
            this.txtPPTFilePath.Location = new System.Drawing.Point(183, 144);
            this.txtPPTFilePath.Name = "txtPPTFilePath";
            this.txtPPTFilePath.Size = new System.Drawing.Size(243, 20);
            this.txtPPTFilePath.TabIndex = 12;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.BackColor = System.Drawing.Color.Transparent;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.ForeColor = System.Drawing.SystemColors.Highlight;
            this.label1.Location = new System.Drawing.Point(67, 111);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(110, 13);
            this.label1.TabIndex = 9;
            this.label1.Text = "Browse Word File:";
            // 
            // txtWordFilePath
            // 
            this.txtWordFilePath.Location = new System.Drawing.Point(183, 108);
            this.txtWordFilePath.Name = "txtWordFilePath";
            this.txtWordFilePath.Size = new System.Drawing.Size(243, 20);
            this.txtWordFilePath.TabIndex = 8;
            // 
            // FormPPTGenerator
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackgroundImage = global::PPTGenerator.Properties.Resources.Background;
            this.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.ClientSize = new System.Drawing.Size(846, 349);
            this.Controls.Add(this.btnBrowsePPTFile);
            this.Controls.Add(this.btnBrowseWordFile);
            this.Controls.Add(this.btnPPTGenerator);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.txtPPTFilePath);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.txtWordFilePath);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "FormPPTGenerator";
            this.Text = "PPTGenerator";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.OpenFileDialog openWordFileDialog;
        private System.Windows.Forms.Button btnBrowsePPTFile;
        private System.Windows.Forms.Button btnBrowseWordFile;
        private System.Windows.Forms.Button btnPPTGenerator;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtPPTFilePath;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtWordFilePath;
    }
}

