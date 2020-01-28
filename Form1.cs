using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;


namespace PPTGenerator
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnBrowseWordFile_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog
            {
                InitialDirectory = @"C:\",
                Title = "Browse Word Files",

                CheckFileExists = true,
                CheckPathExists = true,

                DefaultExt = "docx",
                Filter = "docx files (*.docx)|*.docx",
                FilterIndex = 2,
                RestoreDirectory = true,

                ReadOnlyChecked = true,
                ShowReadOnly = true
            };

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                txtWordFilePath.Text = openFileDialog1.FileName;
            }
        }

        private void btnPPTGenerator_Click(object sender, EventArgs e)
        {
            //Load documents with content controls
            Microsoft.Office.Interop.Word.Application ap = new Microsoft.Office.Interop.Word.Application();
            Microsoft.Office.Interop.PowerPoint.Application pp = new Microsoft.Office.Interop.PowerPoint.Application();
            Document document = ap.Documents.Open(txtWordFilePath.Text);

            
            Microsoft.Office.Interop.PowerPoint.Presentation objShow;
            var pres = pp.Presentations;
            objShow = pres.Open(@"F:\Kiaan Software\Freelancer\GrahamH\PPTGenerator\Documents\Closing Meeting BLANK PowerPoint Template (2020-01-25) 02 (1) - Copy.pptx"); //  , 0 , 0,Microsoft.Office.Core.MsoTriState.msoTrue);
            

            foreach (ContentControl cc in document.ContentControls)
            {
                if(cc.Tag == "Ship_Name")
                {
                    
                }
            }
        }
    }
}
