using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DocumentTest
{
    //2.2版本
    public partial class DsoOffice : UserControl
    {
        private AxDSOFramer.AxFramerControl dso = new AxDSOFramer.AxFramerControl();
        public DsoOffice()
        {
            InitializeComponent();


            ((System.ComponentModel.ISupportInitialize)(this.dso)).BeginInit();
            this.Controls.Add(this.dso);
            ((System.ComponentModel.ISupportInitialize)(this.dso)).EndInit();
            dso.Dock = DockStyle.Fill;

            dso.Titlebar = false;
            dso.Menubar = false;
            dso.Toolbars = true;
            dso.set_EnableFileCommand(DSOFramer.dsoFileCommandType.dsoFileSave, false);
            dso.set_EnableFileCommand(DSOFramer.dsoFileCommandType.dsoFileSaveAs, false);            
            dso.BackColor = Color.Black;
        }

        private string GetFileType(string fileExtension)
        {
            try
            {
                string sOpenType = "";
                switch (fileExtension.ToLower())
                {
                    case "xls":
                    case "xlsx":
                    case "xlsm":
                    case "xlsb":
                    case "csv":
                        sOpenType = "Excel.Sheet";
                        break;
                    case "doc":
                    case "docx":
                    case "docm":
                    case "rtf":
                        sOpenType = "Word.Document";
                        break;
                    case "ppt":
                    case "pptx":
                    case "pptm":
                        sOpenType = "PowerPoint.Show";
                        break;
                    case "vsd":
                    case "vdx":
                        sOpenType = "Visio.Drawing";
                        break;
                    default:
                        sOpenType = "Word.Document";
                        break;
                }
                return sOpenType;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public void OpenDocument(string filepath)
        {
            string sExt = System.IO.Path.GetExtension(filepath).Replace(".", "");
            dso.Open(filepath, false, GetFileType(sExt), "", "");
        }

        public void SaveDocument()
        {
            try
            {
                this.dso.Save(true, true, null, null);
            }
            catch (Exception ex)
            {
            }
        }
    }
}
