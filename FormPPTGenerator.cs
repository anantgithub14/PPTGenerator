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
using Microsoft.Office.Interop.PowerPoint;
using System.IO;

namespace PPTGenerator
{
    public partial class FormPPTGenerator : Form
    {
        public FormPPTGenerator()
        {
            InitializeComponent();
        }

        private void btnBrowseWordFile_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog
            {
                InitialDirectory = Properties.Settings.Default.DefaultDirectory, 
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
            Microsoft.Office.Interop.Word.Application ap;
            Microsoft.Office.Interop.PowerPoint.Application pp;
            Document document;
            btnPPTGenerator.Enabled = false;


            if (string.IsNullOrEmpty(txtWordFilePath.Text) || string.IsNullOrEmpty(txtPPTFilePath.Text))
            {
                MessageBox.Show("Please select the input files");
            }
            else
            {
                try
                {
                    //Load documents with content controls
                    ap = new Microsoft.Office.Interop.Word.Application();
                    pp = new Microsoft.Office.Interop.PowerPoint.Application();
                    document = ap.Documents.Open(txtWordFilePath.Text);


                    Microsoft.Office.Interop.PowerPoint.Presentation objShow;
                    var pres = pp.Presentations;
                    objShow = pres.Open(txtPPTFilePath.Text); 

                    foreach (ContentControl cc in document.ContentControls)
                    {
                        foreach (Slide slide in objShow.Slides)
                        {
                            foreach (Microsoft.Office.Interop.PowerPoint.Shape shape in slide.Shapes)
                            {
                                if (shape.TextFrame.TextRange.Text.Trim() == cc.Tag.Trim())
                                {
                                    shape.TextFrame.TextRange.Text = cc.XMLMapping.CustomXMLNode.FirstChild.NodeValue;
                                    break;
                                }
                            }
                        }                        
                    }
                    document.Close();
                    FileInfo fi = new FileInfo(txtPPTFilePath.Text);
                    Random random = new Random();
                    string newFileName = fi.Name + "_" + random.Next().ToString()  + ".pptx";
                    objShow.SaveAs(fi.DirectoryName + "\\" + newFileName);
                    objShow.Close();
                    
                    MessageBox.Show("The PPT Generation is completed successfully");
                    btnPPTGenerator.Enabled = true;
                }
                catch(Exception ex)
                {
                    throw (ex);
                    
                }
                finally
                {

                }
            }
        }

        private void btnBrowsePPTFile_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog
            {
                InitialDirectory = @"C:\",
                Title = "Browse PPT File Template",

                CheckFileExists = true,
                CheckPathExists = true,

                DefaultExt = "PPTX",
                Filter = "pptx files (*.pptx)|*.pptx",
                FilterIndex = 2,
                RestoreDirectory = true,

                ReadOnlyChecked = true,
                ShowReadOnly = true
            };

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                txtPPTFilePath.Text = openFileDialog1.FileName;
            }

        }
    }
}
