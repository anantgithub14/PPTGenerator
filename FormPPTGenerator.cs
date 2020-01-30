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
                            objShow.Slides.AddSlide(i, layout);

                            foreach (Microsoft.Office.Interop.PowerPoint.Shape shape in objShow.Slides[i].Shapes)
                            {
                                foreach (ContentControl cc in document.ContentControls)
                                {
                                    if (shape.TextFrame.TextRange.Text.Trim() == cc.Tag.Trim())
                                    {
                                        shape.TextFrame.TextRange.Text = cc.XMLMapping.CustomXMLNode.FirstChild.NodeValue;
                                        break;
                                    }
                                }
                            }
                            
                            i++;
                            break;
                        }
                    }


                    foreach (ContentControl cc in document.ContentControls)
                    {
                        switch(cc.Tag.Trim())
                        {
                            case "Section_Name_01":
                                foreach (CustomLayout layout in objShow.SlideMaster.CustomLayouts)
                                {
                                    if (layout.Name.Equals("Section_Page_01"))
                                    {
                                        objShow.Slides.AddSlide(i, layout);

                                        objShow.Slides[i].Shapes[1].TextFrame.TextRange.Text = cc.Range.Text;
                                        i++;
                                        break;
                                    }
                                }
                                
                                break;
                            case "Section_Name_02":
                                foreach (CustomLayout layout in objShow.SlideMaster.CustomLayouts)
                                {
                                    if (layout.Name.Equals("Section_Page_02"))
                                    {
                                        objShow.Slides.AddSlide(i, layout);
                                        objShow.Slides[i].Shapes[1].TextFrame.TextRange.Text = cc.Range.Text;
                                        i++;
                                        break;
                                    }
                                }
                                break;
                            case "Section_Name_03":
                                foreach (CustomLayout layout in objShow.SlideMaster.CustomLayouts)
                                {
                                    if (layout.Name.Equals("Section_Page_03"))
                                    {
                                        objShow.Slides.AddSlide(i, layout);
                                        objShow.Slides[i].Shapes[1].TextFrame.TextRange.Text = cc.Range.Text;
                                        i++;
                                        break;
                                    }
                                }
                                break;
                            case "Section_Name_04":
                                foreach (CustomLayout layout in objShow.SlideMaster.CustomLayouts)
                                {
                                    if (layout.Name.Equals("Section_Page_04"))
                                    {
                                        objShow.Slides.AddSlide(i, layout);
                                        objShow.Slides[i].Shapes[1].TextFrame.TextRange.Text = cc.Range.Text;
                                        i++;
                                        break;
                                    }
                                }
                                break;
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
                            
                            if(Sections.repeat == "R" && Sections.pictureType == "L")
                            {
                                foreach (CustomLayout layout in objShow.SlideMaster.CustomLayouts)
                                {
                                    if (layout.Name.Equals("Picture Details Landscape Repeat"))
                                    {
                                        objShow.Slides.AddSlide(i, layout);                                        
                                        i++;
                                        break;
                                    }
                                }                                
                            }
                            else if(Sections.repeat == "R" && Sections.pictureType == "P")
                            {
                                foreach (CustomLayout layout in objShow.SlideMaster.CustomLayouts)
                                {
                                    if (layout.Name.Equals("Picture Details Portrait Repeat"))
                                    {
                                        objShow.Slides.AddSlide(i, layout);
                                        i++;
                                        break;
                                    }
                                }                                
                            }
                            else if (Sections.repeat == "R" && Sections.pictureType == "NP")
                            {
                                foreach (CustomLayout layout in objShow.SlideMaster.CustomLayouts)
                                {
                                    if (layout.Name.Equals("Picture Details Landscape NP Repeat"))
                                    {
                                        objShow.Slides.AddSlide(i, layout);
                                        i++;
                                        break;
                                    }
                                }
                            }
                            if (Sections.repeat == "N/A" && Sections.pictureType == "L")
                            {
                                foreach (CustomLayout layout in objShow.SlideMaster.CustomLayouts)
                                {
                                    if (layout.Name.Equals("Picture Details Landscape"))
                                    {
                                        objShow.Slides.AddSlide(i, layout);

                                        foreach(Microsoft.Office.Interop.PowerPoint.Shape shape in objShow.Slides[i].Shapes)
                                        {
                                            if (shape.Title == "Header") shape.TextFrame.TextRange.Text = Sections.pictureHeading;
                                            if(shape.Title == "Left") shape.TextFrame.TextRange.Text = Sections.explanatoryText;
                                            
                                        }

                                        i++;
                                        break;
                                    }
                                }
                            }
                            else if (Sections.repeat == "N/A" && Sections.pictureType == "P")
                            {
                                foreach (CustomLayout layout in objShow.SlideMaster.CustomLayouts)
                                {
                                    if (layout.Name.Equals("Picture Details Portrait"))
                                    {
                                        objShow.Slides.AddSlide(i, layout);
                                        i++;
                                        break;
                                    }
                                }
                            }
                            else if (Sections.repeat == "N/A" && Sections.pictureType == "NP")
                            {
                                foreach (CustomLayout layout in objShow.SlideMaster.CustomLayouts)
                                {
                                    if (layout.Name.Equals("Picture Details Landscape NP"))
                                    {
                                        objShow.Slides.AddSlide(i, layout);
                                        i++;
                                        break;
                                    }
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
