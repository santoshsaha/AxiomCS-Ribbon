using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Data;
using Word = Microsoft.Office.Interop.Word;
using Telerik.Windows.Controls;
using System.ComponentModel;

using AxiomIRISRibbon.Core;
using System.IO;


namespace AxiomIRISRibbon.SForceEdit
{
    /// <summary>
    /// Interaction logic for SForceEditSideBar.xaml
    /// New File added by PES
    /// </summary>
    public partial class CompareSideBar : UserControl
    {

        private Data _d;
        static Microsoft.Office.Interop.Word.Application app;
        static Word.Document _source;
        private string _fileName;


        public CompareSideBar()
        {
           InitializeComponent();
           AxiomIRISRibbon.Utility.setTheme(this);

           _d = Globals.ThisAddIn.getData();

           LoadTemplatesDLL();
        
        }
        public void Create(string filename)
        {
            _fileName = filename;
        
        }
        private void LoadTemplatesDLL()
        {
            try
            {

                DataReturn dr = AxiomIRISRibbon.Utility.HandleData(_d.GetTemplates(true));
                if (!dr.success) return;

                DataTable dt = dr.dt;
                cbTemplates.Items.Clear();

                RadComboBoxItem i;

                // RadComboBoxItem selected = null;
                foreach (DataRow r in dt.Rows)
                {
                    i = new RadComboBoxItem();
                    i.Tag = r["Id"].ToString() ;
                    i.Content = r["Name"].ToString();
                    this.cbTemplates.Items.Add(i);

                }

            }
            catch (Exception ex)
            {
                Logger.Log(ex, "Clone");
            }
        }


        private void btnCompare_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                object missing = System.Reflection.Missing.Value;

                if (this.cbTemplates.SelectedItem != null)
                {

                    string TemplateId = ((RadComboBoxItem)(this.cbTemplates.SelectedItem)).Tag.ToString();
                    string TemplateName = ((RadComboBoxItem)(this.cbTemplates.SelectedItem)).Content.ToString();
                   // Microsoft.Office.Interop.Word.Document tempDoc;

                    DataReturn drAttachemnts = AxiomIRISRibbon.Utility.HandleData(_d.GetTemplateAttach(TemplateId));
                    if (!drAttachemnts.success) return;

                     
                    DataTable dtAttachments = drAttachemnts.dt;
                    string file2name = "";
                    foreach (DataRow rw in dtAttachments.Rows)
                    {
                        byte[] toBytes = Convert.FromBase64String(rw["body"].ToString());
                        file2name = _d.GetTempFilePath(rw["Id"].ToString() + "_" + rw["Name"].ToString());
                        File.WriteAllBytes(file2name, toBytes);
                     //   _source = app.Documents.Open(filename);
                        
                        
                    }
                ////    object o = (object)file2name;
                ////    tempDoc = app.Documents.Open(_fileName);
                //    Word.Document doc1 = Globals.ThisAddIn.Application.Documents.Open(_fileName);
                //   //object doc2 = Globals.ThisAddIn.Application.Documents.Open(file2name);
                //   //doc1.Windows.CompareSideBySideWith(ref doc2);
                //    doc1.Compare(file2name, missing, Microsoft.Office.Interop.Word.WdCompareTarget.wdCompareTargetCurrent, true, false, false, false, false);

                // //   doc1.Windows.CompareSideBySideWith(ref o);
                //    // MessageBox.Show(TemplateId);



                    Microsoft.Office.Interop.Word.Document tempDoc1;
                    Microsoft.Office.Interop.Word.Document tempDoc2;
                    Microsoft.Office.Interop.Word.Application app = Globals.ThisAddIn.Application;



                    object newFilenameObject2 = file2name;
                    tempDoc2 = app.Documents.Open(ref newFilenameObject2, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);

                    object newFilenameObject1 = _fileName;
                    tempDoc1 = app.Documents.Open(ref newFilenameObject1, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                   ref missing, ref missing);

                    object o = tempDoc2;
                    tempDoc1.Windows.CompareSideBySideWith(ref o);

                }
                else
                {
                    MessageBox.Show("Select a template");

                }
            }
            catch (Exception ex)
            {
                Logger.Log(ex, "Clone");
            }
        }


    }
}
