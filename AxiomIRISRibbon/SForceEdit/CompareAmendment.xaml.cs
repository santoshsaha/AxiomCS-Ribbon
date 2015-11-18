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
using Telerik.Windows.Controls;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using System.Data;
using AxiomIRISRibbon.Core;
using System.IO;

namespace AxiomIRISRibbon.SForceEdit
{
    /// <summary>
    /// Interaction logic for CompareAmendment.xaml
    ///    NEW File Added by PES
    /// </summary>
    public partial class CompareAmendment : RadWindow
    {
        private Data _d;
        private string _attachmentid;
        private string _versionid;
        private double _versionNumber;
        private string _strToAgreementId; 
        private string _strTemplate;
        private string _strAttachmentName;
        private DataRow _allDr;
        RadComboBoxItem selected = null;
        public CompareAmendment()
        {
            InitializeComponent();
            AxiomIRISRibbon.Utility.setTheme(this);

            _d = Globals.ThisAddIn.getData();
         

        }
        public void Create(string attachmentid, string versionid, string attachmentName)
        {


            _attachmentid = attachmentid;
            _versionid = versionid;
            _strAttachmentName = attachmentName;
            LoadTemplatesDLL();
          
        }
        private void chkMaster_checked(object sender, RoutedEventArgs e)
        {         
           this.radComboAmendment.SelectedIndex = 0;

        }
        private void cbAmendment_SelectionChanged(object sender, RoutedEventArgs e)
        {

            this.chkMaster.IsChecked = false;

        }
        private void btnOpen_Click(object sender, RoutedEventArgs e)
        {

            if (this.radComboAmendment.SelectedItem == null)
            {
                MessageBox.Show("Select one document");
            }
            else
            {
                double maxId = Convert.ToDouble(_versionNumber + 1);
                string VersionName = "Version " + (maxId).ToString();
                string VersionNumber = maxId.ToString();

                // Create Version 0 or lower version in To
                DataReturn drCreate = AxiomIRISRibbon.Utility.HandleData(_d.CreateVersion("", _strToAgreementId, _strTemplate, VersionName, VersionNumber, _allDr));
                string newVersionId = drCreate.id;

                //Create attachments in To
                DataReturn drVersionAttachemnts = AxiomIRISRibbon.Utility.HandleData(_d.GetVersionAllAttachments(_versionid));
                if (!drVersionAttachemnts.success) return;
                DataTable dtAttachments = drVersionAttachemnts.dt;

                if (dtAttachments.Rows.Count == 0)
                {
                    MessageBox.Show("Attachments not avilable in source Version");
                }
                else
                {
                    string filename = "";
                    foreach (DataRow rw in dtAttachments.Rows)
                    {
                        filename = rw["Name"].ToString();
                        string body = rw["body"].ToString();
                        _d.saveAttachmentstoSF(newVersionId, filename, body);
                    }

                    //Get Attachments
                    DataReturn drAttachemnts = AxiomIRISRibbon.Utility.HandleData(_d.GetAllAttachments(newVersionId));
                    if (!drAttachemnts.success) return;
                    DataTable dtAllAttachments = drAttachemnts.dt;

                    //Open attachment with compare screeen
                    OpenAttachment(dtAllAttachments, newVersionId, _strToAgreementId, _strTemplate, VersionName, VersionNumber);
                    Globals.Ribbons.Ribbon1.CloseWindows();
                    this.Close();
                }

            }

        }
   private void LoadTemplatesDLL()
        {   try
            {

            
              //  DataReturn drTemplateForVersion = AxiomIRISRibbon.Utility.HandleData(_d.GetTemplateForVersion(_versionid));
                 DataReturn drTemplateForVersion = AxiomIRISRibbon.Utility.HandleData(_d.GetAgreementsForVersion("",_versionid));
                _allDr = drTemplateForVersion.dt.Rows[0];
                _versionNumber = Convert.ToDouble(drTemplateForVersion.dt.Rows[0]["version_number__c"]);
                _strToAgreementId = drTemplateForVersion.dt.Rows[0]["matter__c"].ToString();
                _strTemplate = drTemplateForVersion.dt.Rows[0]["Template__c"].ToString();
                DataReturn dr = AxiomIRISRibbon.Utility.HandleData(_d.GetamendmentTemplate(_strTemplate));


             //   DataReturn dr = AxiomIRISRibbon.Utility.HandleData(_d.GetTemplates(true));
                if (!dr.success) return;

                DataTable dt = dr.dt;

             
                RadComboBoxItem i;

                i = new RadComboBoxItem();
                i.Tag = "select";
                i.Content = "";
                this.radComboAmendment.Items.Add(i);
                selected = i;
                foreach (DataRow r in dt.Rows)
                {
                    i = new RadComboBoxItem();
                    i.Tag = r["Id"].ToString();
                    i.Content = r["Name"].ToString();
                    this.radComboAmendment.Items.Add(i);

                }
                this.radComboAmendment.SelectedItem = "select";
          
            }
            catch (Exception ex)
            {
                Logger.Log(ex, "Clone");
            }
        }

    private  void OpenAttachment(DataTable dt,string versionid, string matterid, string templateid, string versionName, string versionNumber)
        {
            try
            {
                object missing = System.Reflection.Missing.Value;

                string TemplateId = ((RadComboBoxItem)(this.radComboAmendment.SelectedItem)).Tag.ToString();
                DataReturn drAttachemnts = AxiomIRISRibbon.Utility.HandleData(_d.GetTemplateAttach(TemplateId));
                if (!drAttachemnts.success) return;
                DataTable dtAttachments = drAttachemnts.dt;
                string file2name = "";
                foreach (DataRow rw in dtAttachments.Rows)
                {
                    byte[] toBytes = Convert.FromBase64String(rw["body"].ToString());
                    file2name = _d.GetTempFilePath(rw["Id"].ToString() + "_" + rw["Name"].ToString());
                    File.WriteAllBytes(file2name, toBytes);


                }      
                    var res = from row in dt.AsEnumerable()
                              where 
                              (row.Field<string>("Name").Contains(".doc") ||
                              row.Field<string>("ContentType").Contains("msword"))
                              select row;
                    if (res.Count() > 1)
                    {

                      foreach (DataRow rw in dt.Rows)
                        {
                            if( (rw["Name"].ToString().Contains(".doc"))&&(_strAttachmentName==rw["Name"].ToString()))
                            {
                                byte[] toBytes = Convert.FromBase64String(rw["body"].ToString());
                                string filename = _d.GetTempFilePath(rw["Id"].ToString() + "_" + rw["Name"].ToString());

                                File.WriteAllBytes(filename, toBytes);
                                // _source = app.Documents.Open(filename);


                                Microsoft.Office.Interop.Word.Document tempDoc1;
                                Microsoft.Office.Interop.Word.Document tempDoc2;
                                Microsoft.Office.Interop.Word.Application app = Globals.ThisAddIn.Application;



                                object newFilenameObject2 = file2name;
                                tempDoc2 = app.Documents.Open(ref newFilenameObject2, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                                ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);

                                object newFilenameObject1 = filename;
                                tempDoc1 = app.Documents.Open(ref newFilenameObject1, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                                ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                               ref missing, ref missing);

                                object o = tempDoc2;
                                tempDoc1.Windows.CompareSideBySideWith(ref o);
                                Globals.Ribbons.Ribbon1.CloseWindows();
                            }
                        }
                    }
            }
            catch (Exception ex) { Logger.Log(ex, "OpenAttachment"); }
           
        }


    
    }
}
