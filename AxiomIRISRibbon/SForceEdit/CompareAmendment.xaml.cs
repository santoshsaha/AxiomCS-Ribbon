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
        private string _newVersionId;
        private string _versionName;
        private double _versionNumber;
        private string _strToAgreementId;
        private string _strTemplate;
        private string _strSelectedAttachmentName;
        private DataRow _allDr;
        RadComboBoxItem selected = null;
        private Word.Document _doc;
        private string _fileToSaveAsAgreement;
        private static string _RightFilePath;
        private string _LeftFilePath;

        static Microsoft.Office.Interop.Word.Document tempDoc1;
       static  Microsoft.Office.Interop.Word.Document tempDoc2;

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
            _strSelectedAttachmentName = attachmentName;
            LoadTemplatesDLL();

        }
        private void LoadTemplatesDLL()
        {
            try
            {
                DataReturn drTemplateForVersion = AxiomIRISRibbon.Utility.HandleData(_d.GetAgreementsForVersion("", _versionid));
                _allDr = drTemplateForVersion.dt.Rows[0];
                _versionNumber = Convert.ToDouble(drTemplateForVersion.dt.Rows[0]["version_number__c"]);
                _strToAgreementId = drTemplateForVersion.dt.Rows[0]["matter__c"].ToString();
                _strTemplate = drTemplateForVersion.dt.Rows[0]["Template__c"].ToString();
                DataReturn dr = AxiomIRISRibbon.Utility.HandleData(_d.GetAmendmentTemplate(_strTemplate,false));
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
                //Logger.Log(ex, "Clone");
            }
        }
        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
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

            if (this.radComboAmendment.SelectedItem == null && this.chkMaster.IsChecked == false)
            {
                MessageBox.Show("Select any template from dropdown  or master checkbox");
            }
            else
            {
                string strTemplateId = string.Empty;
                if (this.radComboAmendment.SelectedItem != null)
                {
                    strTemplateId = ((RadComboBoxItem)(this.radComboAmendment.SelectedItem)).Tag.ToString();
                    if ((strTemplateId == "selected") && (chkMaster.IsChecked == true))
                    {
                        DataReturn dr = AxiomIRISRibbon.Utility.HandleData(_d.GetAmendmentTemplate(_strTemplate, true));
                        strTemplateId = dr.dt.Rows[0]["Id"].ToString();
                    }
                }

                double maxId = Convert.ToDouble(_versionNumber + 1);
                string VersionName = "Version " + (maxId).ToString();
                _versionName = VersionName;
                string VersionNumber = maxId.ToString();
              

                // Create Version 0 or lower version in To
                DataReturn drCreate = AxiomIRISRibbon.Utility.HandleData(_d.CreateVersion("", _strToAgreementId, _strTemplate, VersionName, VersionNumber, _allDr));
                _newVersionId = drCreate.id;
               
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
                     //   if (filename != _strSelectedAttachmentName)
                    //    {
                            string body = rw["body"].ToString();
                            _d.saveAttachmentstoSF(_newVersionId, filename, body);
                    //    }
                    }

                    //Save template into version as amendtment
                    DataReturn drTemplate = AxiomIRISRibbon.Utility.HandleData(_d.GetTemplateAttach(strTemplateId));
                    if (!drTemplate.success) return;
                    DataTable dtTemplate = drTemplate.dt;
                    string fileNameTemplate = VersionName + "_Amendment.docx";
                    _d.saveAttachmentstoSF(_newVersionId, fileNameTemplate, dtTemplate.Rows[0]["body"].ToString());

                    //Get Attachments
                    DataReturn drAttachemnts = AxiomIRISRibbon.Utility.HandleData(_d.GetAllAttachments(_newVersionId));
                    if (!drAttachemnts.success) return;
                    DataTable dtAllAttachments = drAttachemnts.dt;

                    //Open attachment with compare screeen
                    PrepareFiles(dtAllAttachments, _newVersionId, _strToAgreementId, _strTemplate, VersionName, VersionNumber, fileNameTemplate);
                    Globals.Ribbons.Ribbon1.CloseWindows();
                    this.Close();
                }

            }

        }


        private void PrepareFiles(DataTable dt, string versionid, string matterid, string templateid, string versionName, string versionNumber, string strFileNameTemplate)
        {
            try
            {

                string fileAmendmentDocument = string.Empty, fileAmendmentTemplate = string.Empty, newAttachmentId = string.Empty;
                var res = from row in dt.AsEnumerable()
                          where
                          (row.Field<string>("Name").Contains(".doc") ||
                          row.Field<string>("ContentType").Contains("msword"))
                          select row;
                if (res.Count() > 1)
                {

                    foreach (DataRow rw in dt.Rows)
                    {
                        if (rw["Name"].ToString().Contains(".doc"))
                        {

                            if (rw["Name"].ToString() == _strSelectedAttachmentName)
                            {
                                byte[] toBytes = Convert.FromBase64String(rw["body"].ToString());
                               // fileAmendmentDocument = _d.GetTempFilePath(rw["Id"].ToString() + "_" + rw["Name"].ToString());
                                fileAmendmentDocument = _d.GetTempFilePath(rw["Id"].ToString() + _strSelectedAttachmentName);
                                File.WriteAllBytes(fileAmendmentDocument, toBytes);
                                newAttachmentId = rw["Id"].ToString();
                            }
                            else if (rw["Name"].ToString() == strFileNameTemplate)
                            {
                                byte[] toBytes = Convert.FromBase64String(rw["body"].ToString());
                              //  fileAmendmentTemplate = _d.GetTempFilePath(rw["Id"].ToString() + "_AmendmentTemplate");
                                fileAmendmentTemplate = _d.GetTempFilePath(rw["Id"].ToString() + "_" + rw["Name"].ToString());
                                File.WriteAllBytes(fileAmendmentTemplate, toBytes);
                                _fileToSaveAsAgreement = fileAmendmentTemplate;
                             
                            }
                            // _source = app.Documents.Open(filename);
                        }
                    }
                    if (fileAmendmentDocument == string.Empty && fileAmendmentTemplate == string.Empty)
                    {
                        MessageBox.Show("Files not avilable");
                    }
                    else
                    {
                        CombineDocs(fileAmendmentDocument, fileAmendmentTemplate, newAttachmentId);

                        OpenFiles();
                    }
                }
            }
            catch (Exception ex)
            { //Logger.Log(ex, "OpenAttachment");
            }

        }

        private void OpenFiles()
        {
            object missing = System.Reflection.Missing.Value;
            string fileAmendmentDocument = string.Empty, fileAmendmentTemplate = string.Empty;
            string vfilename = _versionName.Replace(" ", "_");
            string fileNameTemplate = _versionName + "_Amendment";
            DataReturn drAttachemnts = AxiomIRISRibbon.Utility.HandleData(_d.GetAllAttachments(_newVersionId));
            if (!drAttachemnts.success) return;
            DataTable dtAllAttachments = drAttachemnts.dt;
            var res = from row in dtAllAttachments.AsEnumerable()
                      where
                      (row.Field<string>("Name").Contains(".doc") ||
                      row.Field<string>("ContentType").Contains("msword"))
                      select row;
            if (res.Count() > 1)
            {

                foreach (DataRow rw in dtAllAttachments.Rows)
                {
                    if (rw["Name"].ToString().Contains(".doc"))
                    {

                        if (rw["Name"].ToString() == vfilename + ".docx")
                        {
                            byte[] toBytes = Convert.FromBase64String(rw["body"].ToString());
                            fileAmendmentDocument = _d.GetTempFilePath(rw["Id"].ToString() + "_" + vfilename);
                            File.WriteAllBytes(fileAmendmentDocument, toBytes);
                        }
                        else if (rw["Name"].ToString() == fileNameTemplate + ".docx")
                        {
                            byte[] toBytes = Convert.FromBase64String(rw["body"].ToString());
                            fileAmendmentTemplate = _d.GetTempFilePath(rw["Id"].ToString() + "_" + fileNameTemplate);
                            File.WriteAllBytes(fileAmendmentTemplate, toBytes);

                        }
                    }
                }
                CompareSideBySide(fileAmendmentDocument, fileAmendmentTemplate);
              /*  Microsoft.Office.Interop.Word.Document tempDoc1;
                Microsoft.Office.Interop.Word.Document tempDoc2;
                Microsoft.Office.Interop.Word.Application appl = Globals.ThisAddIn.Application;

                object newFilenameObject1 = fileAmendmentDocument;
                tempDoc1 = appl.Documents.Open(ref newFilenameObject1, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
               ref missing, ref missing);

                object newFilenameObject2 = fileAmendmentTemplate;
                tempDoc2 = appl.Documents.Open(ref newFilenameObject2, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);



                object o = tempDoc2;
                tempDoc1.Windows.CompareSideBySideWith(ref o);
                */
            }
        }

        public void CompareSideBySide(string fileAmendmentDocumentPath, string fileAmendmentTemplatePath)
        {
            try
            {
                _RightFilePath = fileAmendmentTemplatePath;
                _LeftFilePath = fileAmendmentDocumentPath;

                object start = 0;
                object end = 0;
                object missing = System.Reflection.Missing.Value;

             
                Microsoft.Office.Interop.Word.Application app = Globals.ThisAddIn.Application;

                object newFilenameObject1 = fileAmendmentTemplatePath;
                tempDoc1 = app.Documents.Open(ref newFilenameObject1, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing);

                object newFilenameObject2 = fileAmendmentDocumentPath;
                tempDoc2 = app.Documents.Open(ref newFilenameObject2, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);


                object o = tempDoc1;
                tempDoc2.Windows.CompareSideBySideWith(ref o);
                MessageBox.Show("Process complete");
                tempDoc2.AcceptAllRevisions();
                tempDoc2.TrackRevisions = true;

                /*  foreach (Microsoft.Office.Interop.Word.Section s in tempDoc2.Sections)
                  {
                      foreach (Microsoft.Office.Interop.Word.Revision r in s.Range.Revisions)
                      {
                          counter += r.Range.Words.Count;
                          if (r.Type == Microsoft.Office.Interop.Word.WdRevisionType.wdRevisionDelete) // Deleted
                              delcnt += r.Range.Words.Count;
                          if (r.Type == Microsoft.Office.Interop.Word.WdRevisionType.wdRevisionInsert) // Inserted
                              inscnt += r.Range.Words.Count;
                          wr = r.Range.Text;
                          //  r.Range.AutoFormat();
                          // wr = r.Range.AutoFormat();
                          Microsoft.Office.Interop.Word.Range rng = tempDoc1.Range(0, 0);
                          rng.Text = wr;
                          rng.Bold = 1;
                          if (r.Type == Microsoft.Office.Interop.Word.WdRevisionType.wdRevisionProperty) // Formatting (bold,italics)
                              inscnt += r.Range.Words.Count;
                          //object o = tempDoc1;
                          //_source1.Windows.CompareSideBySideWith(ref o);

                      }
                  } */
            }

            catch (Exception exe)
            {
                MessageBox.Show(exe.ToString());
            }

        }
        private void CombineDocs(string fileAmendmentDocumentPath, string fileAmendmentTemplatePath, string newAttachmentId)
        {
            Microsoft.Office.Interop.Word.Application app = Globals.ThisAddIn.Application;
            object missing = System.Reflection.Missing.Value;
            Microsoft.Office.Interop.Word.Document tempAmendmentTemplate;
            object objAmendmentTemplate = fileAmendmentTemplatePath;
            tempAmendmentTemplate = app.Documents.Open(ref objAmendmentTemplate, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                          ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                         ref missing, ref missing);


            Microsoft.Office.Interop.Word.Fields fs = tempAmendmentTemplate.Fields;
            foreach (Microsoft.Office.Interop.Word.Field f in fs)
            {
                f.Select();
                tempAmendmentTemplate.Application.Selection.InsertFile(fileAmendmentDocumentPath);
            }
        
            //string vfilename = _versionName.Replace(" ", "_") + ".docx";
            DataReturn dr = SaveContract(newAttachmentId, fileAmendmentTemplatePath);
            app.Documents.Close();
          /*  foreach (DataRow rw in dr.dt.Rows)
            {

                byte[] toBytes = Convert.FromBase64String(rw["body"].ToString());
                fileAmendmentDocumentPath = _d.GetTempFilePath(rw["Id"].ToString() + "_" + rw["Name"].ToString());
                File.WriteAllBytes(fileAmendmentDocumentPath, toBytes);
            }*/
            //   SaveContract(false, true);

          
        }

        public static void TrackDocument()
        {

            object missing = System.Reflection.Missing.Value;
            Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.Application();
            Object destFile = _RightFilePath;
            object start = 0;
            object end = 0;
            int counter = 0;
            int delcnt = 0;
            int inscnt = 0;
            String wr = null;
            // Document leftDoc = app.Documents.Add(templateDoc);
            // Document leftDoc = app.Documents.Open(ref destFile, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
            //ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
            //ref missing, ref missing);
            // tempDoc2.AcceptAllRevisions();
            foreach (Microsoft.Office.Interop.Word.Section s in tempDoc2.Sections)
            {
                for (int rnumber = s.Range.Revisions.Count; rnumber > 0; rnumber--)
                {


                    //foreach (Microsoft.Office.Interop.Word.Revision r in s.Range.Revisions)
                    //{
                    Microsoft.Office.Interop.Word.Revision r = s.Range.Revisions[rnumber];
                    counter += r.Range.Sections.Count;
                    counter += s.Range.Sections.Count;
                    if (r.Type == Microsoft.Office.Interop.Word.WdRevisionType.wdRevisionDelete) // Deleted
                    {
                        delcnt += r.Range.Words.Count;
                        wr = r.Range.Text;
                        Microsoft.Office.Interop.Word.Range rng = tempDoc1.Range(0, 0);
                        rng.Text = wr;
                    }
                    if (r.Type == Microsoft.Office.Interop.Word.WdRevisionType.wdRevisionInsert) // Inserted
                    {
                        inscnt += r.Range.Words.Count;
                        wr = r.Range.Text;

                        //  r.Range.AutoFormat();
                        // wr = r.Range.AutoFormat();
                        Microsoft.Office.Interop.Word.Range rng = tempDoc1.Range(0, 0);
                        rng.Text = wr;
                        rng.Bold = 1;
                    }
                    if (r.Type == Microsoft.Office.Interop.Word.WdRevisionType.wdRevisionProperty) // Formatting (bold,italics)
                    {
                        inscnt += r.Range.Words.Count;
                        Microsoft.Office.Interop.Word.Range rng = tempDoc1.Range(0, 0);
                        rng.Text = wr;
                        //object o = tempDoc1;
                        //_source1.Windows.CompareSideBySideWith(ref o);
                    }
                }
            }
        }


        public DataReturn SaveContract(string newAttachmentId, string fileAmendmentTemplatePath)
        {
            string strFileAttached = fileAmendmentTemplatePath;
            //Save the Contract    
          
            DataReturn dr;
            _doc = Globals.ThisAddIn.Application.ActiveDocument;      
              
                _doc.SaveAs2(FileName: strFileAttached, FileFormat: Word.WdSaveFormat.wdFormatXMLDocument, CompatibilityMode: Word.WdCompatibilityMode.wdCurrent);           
                string filenamecopy = AxiomIRISRibbon.Utility.SaveTempFile(_versionid + "X");
                Word.Document dcopy = Globals.ThisAddIn.Application.Documents.Add(strFileAttached, Visible: false);
                dcopy.SaveAs2(FileName: filenamecopy, FileFormat: Word.WdSaveFormat.wdFormatXMLDocument, CompatibilityMode: Word.WdCompatibilityMode.wdCurrent);

                var docclose = (Microsoft.Office.Interop.Word._Document)dcopy;
                docclose.Close();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(docclose);

              
                string vfilename = _versionName.Replace(" ", "_") + ".docx";
                dr = AxiomIRISRibbon.Utility.HandleData(_d.UpdateFile(newAttachmentId, vfilename, filenamecopy));

                return dr;
        }
        public bool SaveContract(bool ForceSave, bool SaveDoc)
        {
            string strFileAttached = _fileToSaveAsAgreement;
            //Save the Contract    
            Globals.ThisAddIn.RemoveSaveHandler(); // remove the save handler to stop the save calling the save etc.

            Globals.ThisAddIn.ProcessingStart("Save Contract");
            DataReturn dr;
            _doc = Globals.ThisAddIn.Application.ActiveDocument;

            dr = AxiomIRISRibbon.Utility.HandleData(_d.SaveVersion(_versionid, "", _strTemplate, _versionName, (_versionNumber+1).ToString()));
            if (!dr.success) return false;
            _versionid = dr.id;

            if (SaveDoc)
            {

                //Save the file as an attachment
                //save this to a scratch file
                Globals.ThisAddIn.ProcessingUpdate("Save Scratch");
                //   string filename = AxiomIRISRibbon.Utility.SaveTempFile(_versionid);
                _doc.SaveAs2(FileName: strFileAttached, FileFormat: Word.WdSaveFormat.wdFormatXMLDocument, CompatibilityMode: Word.WdCompatibilityMode.wdCurrent);

                //Save a copy!
                Globals.ThisAddIn.ProcessingUpdate("Save Copy");
                string filenamecopy = AxiomIRISRibbon.Utility.SaveTempFile(_versionid + "X");
                Word.Document dcopy = Globals.ThisAddIn.Application.Documents.Add(strFileAttached, Visible: false);
                dcopy.SaveAs2(FileName: filenamecopy, FileFormat: Word.WdSaveFormat.wdFormatXMLDocument, CompatibilityMode: Word.WdCompatibilityMode.wdCurrent);

                var docclose = (Microsoft.Office.Interop.Word._Document)dcopy;
                docclose.Close();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(docclose);

                //Now save the file - change this to always save as the version name

                Globals.ThisAddIn.ProcessingUpdate("Save To SalesForce");
                string vfilename = _versionName.Replace(" ", "_") + ".docx";
                dr = AxiomIRISRibbon.Utility.HandleData(_d.UpdateFile(_attachmentid, vfilename, filenamecopy));

            }
            Globals.ThisAddIn.AddSaveHandler(); // add it back in
            Globals.ThisAddIn.ProcessingStop("End");
            return true;
        }

      
    }
}
