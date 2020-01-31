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
using System.Globalization;

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

        private void btnBrowsePPTFile_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog
            {
                InitialDirectory = Properties.Settings.Default.DefaultDirectory,
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
                    document = ap.Documents.Open(txtWordFilePath.Text,ReadOnly:false );

                    Microsoft.Office.Interop.PowerPoint.Presentation objShow;
                    var pres = pp.Presentations;
                    objShow = pres.Open(txtPPTFilePath.Text);

                    int i = 1;

                    foreach (CustomLayout layout in objShow.SlideMaster.CustomLayouts)
                    {
                        if (layout.Name.Equals("Front Main Page"))
                        {
                            
                            foreach (Microsoft.Office.Interop.PowerPoint.Shape shape in layout.Shapes)
                            {
                                foreach (ContentControl cc in document.ContentControls)
                                {
                                    if (shape.Name.Trim() == cc.Tag.Trim())
                                    {
                                        shape.TextFrame.TextRange.Text = cc.XMLMapping.CustomXMLNode.FirstChild.NodeValue;
                                        break;
                                    }
                                }
                            }

                            objShow.Slides.AddSlide(i, layout);

                            i++;
                            break;
                        }
                    }


                    foreach (ContentControl cc in document.ContentControls)
                    {

                        if(cc.Tag.Contains("Section_Name"))
                        {
                            foreach (CustomLayout layout in objShow.SlideMaster.CustomLayouts)
                            {
                                if (layout.Name.Equals("Section_Page_" + cc.Tag.Trim().Substring(cc.Tag.Trim().Length-2)))
                                {
                                    objShow.Slides.AddSlide(i, layout);

                                    objShow.Slides[i].Shapes[1].TextFrame.TextRange.Text = cc.Range.Text;
                                    i++;
                                    break;
                                }
                            }
                        }

                        
                        if (cc.Tag.Contains("PP_Heading"))
                        {                            
                            Sections.pictureHeading = cc.Range.Text.Trim();
                        }
                        else if (cc.Tag.Contains("PP_Main_Text"))
                        {
                            Sections.explanatoryText = cc.Range.Text.Trim();
                        }
                        else if (cc.Tag.Contains("PP_Repeat_NC"))
                        {
                            Sections.repeat = cc.Range.Text.Trim();
                        }
                        else if (cc.Tag.Contains("PP_Photo_Style"))
                        {
                            Sections.pictureType = cc.Range.Text.Trim();
                        }
                        else if (cc.Tag.Contains("PP_Picture"))
                        {                           

                            foreach (CustomLayout layout in objShow.SlideMaster.CustomLayouts)
                            {
                                if (layout.Name.Equals("Template_" + Sections.repeat + "_" + Sections.pictureType))
                                {
                                    foreach (Microsoft.Office.Interop.PowerPoint.Shape shape in layout.Shapes)
                                    {
                                        if (shape.Name.Contains("Heading")) shape.TextFrame.TextRange.Text = Sections.pictureHeading;
                                        if (shape.Name.Contains("Descriptive")) shape.TextFrame.TextRange.Text = Sections.explanatoryText;
                                        if (shape.Name.Equals("Picture")) {
                                            cc.Copy(); shape.TextFrame.TextRange.Paste(); break;                                            
                                        }
                                      
                                    }

                                    objShow.Slides.AddSlide(i, layout);
                                    i++;
                                    break;
                                }
                            }
                        }

                                    
                    }
                    document.Close();
                    FileInfo fi = new FileInfo(txtPPTFilePath.Text);                    
                    string DateString = System.DateTime.Now.Year.ToString() + " - " + System.DateTime.Now.Month.ToString() + " - " + System.DateTime.Now.Day.ToString();                    
                    string newFileName = Properties.Settings.Default.FileNameFormat + " " + "(" + DateString + ").pptx";
  
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


    }
}
