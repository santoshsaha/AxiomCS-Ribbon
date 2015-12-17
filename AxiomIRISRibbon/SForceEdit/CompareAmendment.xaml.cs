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
using System.Threading;
using System.Windows.Threading;
using System.ComponentModel;


namespace AxiomIRISRibbon.SForceEdit
{
    /// <summary>
    /// Interaction logic for CompareAmendment.xaml
    ///    NEW File Added by PES
    ///    Story2 - On click of amendment button this screen will be avilable
    /// </summary>
    public partial class CompareAmendment : RadWindow
    {
        static Word.Document objtempAmendmentTemplate;
        static Word.Document objtempDocAmendment;

        private static Data _d;
        private static string _attachmentid;
        private static string _versionid;
        private static string _versionName;
        private static double _versionNumber;
        private static string _strTemplate;
        private string _strSelectedAttachmentName;
        private static Word.Document _doc;
        private static string _strAmendmentTemplatePath;
        private static string _strAmendmentDocumentPath;
        private static string _strNewAttachmentId;
        private static string _strAmendmentAttachmentId;
        private string _fileToSaveAsAgreement;
        private static string _newVersionId;
        private string _strToAgreementId;
        private DataRow _allDr;
        private static string _strAmendmentTemplateName;
        private static string _strAmendmentDocumentName;
        RadComboBoxItem selected = null;
        BackgroundWorker bsyCompareAmndIndicatorBackgroundWorker;

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
                DataReturn dr = AxiomIRISRibbon.Utility.HandleData(_d.GetAmendmentTemplate(_strTemplate, false));
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

        void bsyCompareAmndIndicatorBackgroundWorker_DoWork(object sender, DoWorkEventArgs e, string strTemplateId)
        {
            BackgroundWorker worker = sender as BackgroundWorker;
            PerformOpening(worker, e, strTemplateId);
        }
        void bsyCompareAmndIndicatorBackgroundWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {

            bsyCompareAmndIndicatorBackgroundWorker.DoWork -= (obj, ev) => bsyCompareAmndIndicatorBackgroundWorker_DoWork(obj, ev, null);
            bsyCompareAmndIndicatorBackgroundWorker.RunWorkerCompleted -= bsyCompareAmndIndicatorBackgroundWorker_RunWorkerCompleted;

            bsyCompareAmndIndc.IsBusy = false;
            bsyCompareAmndIndc.BusyContent = "";
            Globals.Ribbons.Ribbon1.CloseWindows();
            this.Close();
        }

        protected void PerformOpening(BackgroundWorker worker, DoWorkEventArgs e, string strTemplateId)
        {
            try
            {
                double maxId = Convert.ToDouble(_versionNumber + 1);
                string VersionName = "Version " + (maxId).ToString();
                _versionName = VersionName;
                string VersionNumber = maxId.ToString();


                // Create Version 2 or lower version in To
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
                    string filename = string.Empty, body = string.Empty;
                    foreach (DataRow rw in dtAttachments.Rows)
                    {
                        filename = rw["Name"].ToString();
                        if (filename == _strSelectedAttachmentName)
                        {
                            body = rw["body"].ToString();
                            _d.saveAttachmentstoSF(_newVersionId, VersionName.Replace(" ", "_") + ".docx", body);
                            _strSelectedAttachmentName = VersionName.Replace(" ", "_") + ".docx";
                        }

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

                    //Open attachment with compare view
                    OpenFiles();

                    //Combine docs and compare screeen
                   // CombineFiles(dtAllAttachments, _newVersionId, _strToAgreementId, _strTemplate, VersionName, VersionNumber, fileNameTemplate);
                  
                }
            }
            catch (Exception ex)
            {
                
            }
        }
        private void btnOpen_Click(object sender, RoutedEventArgs e)
        {
            try
            {

                if (this.radComboAmendment.SelectedItem == null && this.chkMaster.IsChecked == false)
                {
                    MessageBox.Show("Please select either one template from dropdown  or select master checkbox");
                }
                else
                {
                    string strTemplateId = string.Empty;
                    if (chkMaster.IsChecked == true)
                    {
                        DataReturn dr = AxiomIRISRibbon.Utility.HandleData(_d.GetAmendmentTemplate("", true));
                        strTemplateId = dr.dt.Rows[0]["Id"].ToString();
                    }
                    else if (this.radComboAmendment.SelectedItem != null)
                    {
                        strTemplateId = ((RadComboBoxItem)(this.radComboAmendment.SelectedItem)).Tag.ToString();


                    }
                    if (strTemplateId == "select")
                    {
                        MessageBox.Show("Please select either one template from dropdown  or select master checkbox");
                    }
                    else
                    {
                        bsyCompareAmndIndc.IsBusy = true;
                        bsyCompareAmndIndc.BusyContent = "Loading ...";
                        bsyCompareAmndIndicatorBackgroundWorker = new BackgroundWorker();

                        bsyCompareAmndIndicatorBackgroundWorker.DoWork += (obj, ev) => bsyCompareAmndIndicatorBackgroundWorker_DoWork(obj, ev, strTemplateId);
                        bsyCompareAmndIndicatorBackgroundWorker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(bsyCompareAmndIndicatorBackgroundWorker_RunWorkerCompleted);
                        bsyCompareAmndIndicatorBackgroundWorker.RunWorkerAsync();

                    }
                }
            }
            catch (Exception ex)
            {
            }
            
        }

    

        private void CombineFiles(DataTable dt, string versionid, string matterid, string templateid, string versionName, string versionNumber, string strFileNameTemplate)
        {
            try
            {
                this.Dispatcher.BeginInvoke(DispatcherPriority.Background,
                new Action(
                delegate
                {

                    string fileAmendmentDocumentPath = string.Empty, fileAmendmentTemplatePath = string.Empty;
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
                                    fileAmendmentDocumentPath = _d.GetTempFilePath(rw["Id"].ToString() + _strSelectedAttachmentName);
                                    File.WriteAllBytes(fileAmendmentDocumentPath, toBytes);
                                    _strNewAttachmentId = rw["Id"].ToString();
                                }
                                else if (rw["Name"].ToString() == strFileNameTemplate)
                                {
                                    byte[] toBytes = Convert.FromBase64String(rw["body"].ToString());
                                    fileAmendmentTemplatePath = _d.GetTempFilePath(rw["Id"].ToString() + "_" + rw["Name"].ToString());
                                    File.WriteAllBytes(fileAmendmentTemplatePath, toBytes);
                                    _fileToSaveAsAgreement = fileAmendmentTemplatePath;
                                    _strAmendmentAttachmentId = rw["Id"].ToString();
                                }
                            }
                        }
                        if (fileAmendmentDocumentPath == string.Empty && fileAmendmentTemplatePath == string.Empty)
                        {
                            MessageBox.Show("Files not avilable");
                        }
                        else
                        {
                            //  CombineDocs : fileAmendmentDocument, fileAmendmentTemplate
                            Word.Application app = Globals.ThisAddIn.Application;
                            object missing = System.Reflection.Missing.Value;
                            Word.Document tempAmendmentTemplate;
                            object objAmendmentTemplate = fileAmendmentTemplatePath;
                            tempAmendmentTemplate = app.Documents.Open(ref objAmendmentTemplate, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                                          ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                                         ref missing, ref missing);


                            for (int i = 1; i <= tempAmendmentTemplate.ContentControls.Count; i++)
                            {
                                tempAmendmentTemplate.ContentControls[i].LockContents = false;
                                tempAmendmentTemplate.ContentControls[i].LockContentControl = false;
                               
                            }
                     
                        
                            Word.Fields fs = tempAmendmentTemplate.Fields;
                            bool isMarkerAvailable = false;
                            bool firstMarkerFound = false;
                            foreach (Word.Field f in fs)
                            {
                                if (f.Type == Word.WdFieldType.wdFieldMergeField)
                                {
                                    isMarkerAvailable = true;
                                    f.Select();
                                    if (!firstMarkerFound)
                                    {
                                        //if (f.Result.Text.ToLower().Contains("axiommarker"))
                                        //{
                                            tempAmendmentTemplate.Application.Selection.InsertFile(fileAmendmentDocumentPath);
                                            firstMarkerFound = true;
                                        //}
                                    }
                                }
                            }
                            if (isMarkerAvailable)
                            {
                                DataReturn dr = SaveCombinedDoc(_strNewAttachmentId, fileAmendmentTemplatePath);
                            }

                      
                            app.Documents.Close();
                            //End Coments.

                            //Get the documents again open Sideby side
                            OpenFiles();
                        }
                    }
                }));
            }
            catch (Exception ex)
            { //Logger.Log(ex, "OpenAttachment");
            }

        }

        private void OpenFiles()
        {
            this.Dispatcher.BeginInvoke(DispatcherPriority.Background,
               new Action(
               delegate
               {
                   string fileAmendmentDocumentPath = string.Empty, fileAmendmentTemplatePath = string.Empty;
                   string vfilename = _versionName.Replace(" ", "_") + ".docx";
                   string fileNameTemplate = _versionName + "_Amendment.docx";
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

                               if (rw["Name"].ToString() == vfilename)
                               {
                                   byte[] toBytes = Convert.FromBase64String(rw["body"].ToString());
                                   fileAmendmentDocumentPath = _d.GetTempFilePath(rw["Id"].ToString() + "_" + vfilename);
                                   File.WriteAllBytes(fileAmendmentDocumentPath, toBytes);

                                   _strNewAttachmentId = rw["Id"].ToString();
                               }
                               else if (rw["Name"].ToString() == fileNameTemplate)
                               {
                                   byte[] toBytes = Convert.FromBase64String(rw["body"].ToString());
                                   fileAmendmentTemplatePath = _d.GetTempFilePath(rw["Id"].ToString() + "_" + fileNameTemplate);
                                   File.WriteAllBytes(fileAmendmentTemplatePath, toBytes);

                                   _strAmendmentAttachmentId = rw["Id"].ToString();
                               }
                           }
                       }


                       _strAmendmentTemplatePath = fileAmendmentTemplatePath;
                       _strAmendmentDocumentPath = fileAmendmentDocumentPath;


                       CompareSideBySide(fileAmendmentDocumentPath, fileAmendmentTemplatePath);
                       //  CompareSplitView(fileAmendmentDocumentPath, fileAmendmentTemplatePath);


                       Globals.Ribbons.Ribbon1.CloseWindows();
                   }
               }));
        }

        private static void CompareSideBySide(string fileAmendmentDocumentPath, string fileAmendmentTemplatePath)
        {
            object missing = System.Reflection.Missing.Value;

            // CompareSideBySide : fileAmendmentDocument, fileAmendmentTemplate

            Word.Application app = Globals.ThisAddIn.Application;

            object newFilenameObject1 = fileAmendmentTemplatePath;
            objtempAmendmentTemplate = app.Documents.Open(ref newFilenameObject1, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
            ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
            ref missing, ref missing);


            object newFilenameObject2 = fileAmendmentDocumentPath;
            objtempDocAmendment = app.Documents.Open(ref newFilenameObject2, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
            ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);

            // To unlock Clauses
             for (int i = 1; i <= objtempDocAmendment.ContentControls.Count; i++)
               {
                   objtempDocAmendment.ContentControls[i].LockContents = false;
                   objtempDocAmendment.ContentControls[i].LockContentControl = false;
               }

               for (int i = 1; i <= objtempAmendmentTemplate.ContentControls.Count; i++)
               {
                   objtempAmendmentTemplate.ContentControls[i].LockContents = false;
                   objtempAmendmentTemplate.ContentControls[i].LockContentControl = false;
               }
              
            //AmendmentTemplate - For Save
            Globals.ThisAddIn.AddDocId(objtempAmendmentTemplate, "Contract", "", "AmendmentTemplate");
            //AmendmentDocument - For Save
            Globals.ThisAddIn.AddDocId(objtempDocAmendment, "Contract", "", "AmendmentDocument");

            object o = objtempAmendmentTemplate;


            // Side by side was not working after added below code. TO DO
            //Remove Markup from template doc
            /*  Word.Fields fields = objtempAmendmentTemplate.Fields;
              foreach (Microsoft.Office.Interop.Word.Field f in fields)
              {
                  f.Select();
                  objtempAmendmentTemplate.Application.Selection.InsertParagraph();

              }*/

            //objtempDocAmendment.Activate();
            objtempDocAmendment.Windows.CompareSideBySideWith(ref o);

            objtempDocAmendment.AcceptAllRevisions();
            
            objtempDocAmendment.TrackRevisions = true;
            objtempDocAmendment.ActiveWindow.View.ShowRevisionsAndComments = false;
            objtempAmendmentTemplate.TrackRevisions = true;
            objtempDocAmendment.Activate();
        }
        /*
        private static void CompareSplitView(string fileAmendmentDocumentPath, string fileAmendmentTemplatePath)
        {
            //Compare Split view

            object missing = System.Reflection.Missing.Value;
            Word.Application app = Globals.ThisAddIn.Application;

            objtempDocAmendment = Globals.ThisAddIn.Application.Documents.Add(fileAmendmentDocumentPath);
            // wordAttachment.TrackRevisions = false;
            //  wordAttachment.ShowRevisions = false;
            //  wordAttachment.AcceptAllRevisions();
            object objTemplate = fileAmendmentTemplatePath;
            objtempAmendmentTemplate = app.Documents.Open(ref objTemplate, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
            ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);

            //Compare
            Globals.ThisAddIn.AddDocId(objtempAmendmentTemplate, "AmendmentDocument", "");
            objtempAmendmentTemplate.ActiveWindow.View.ShowRevisionsAndComments = false;
            // objtempAmendmentTemplate.TrackRevisions = false;
            //  objtempAmendmentTemplate.ShowRevisions = false;
            //   objtempAmendmentTemplate.AcceptAllRevisions();

            //Remove Markup from template doc
            Word.Fields fields = objtempAmendmentTemplate.Fields;
            foreach (Microsoft.Office.Interop.Word.Field f in fields)
            {
                f.Select();
                objtempAmendmentTemplate.Application.Selection.InsertParagraph();

            }

            //  Compare code
            objtempAmendmentTemplate.Compare(fileAmendmentDocumentPath, missing, Word.WdCompareTarget.wdCompareTargetNew, true, true, false, false, false);
            app.ActiveWindow.View.SplitSpecial = Word.WdSpecialPane.wdPaneRevisionsVert;
            app.ActiveWindow.ShowSourceDocuments = Word.WdShowSourceDocuments.wdShowSourceDocumentsOriginal;
            app.ActiveWindow.View.RevisionsFilter.Markup = 0;

            // close the temp files
            // var docTemplateClose = (Word._Document)objtempAmendmentTemplate;
            //  docTemplateClose.Close(SaveChanges: false);
            //   var docAttachmentClose = (Word._Document)objtempDocAmendment;
            //  docAttachmentClose.Close(SaveChanges: false);

            objtempAmendmentTemplate.Activate();

            //End Compare


        }
        */
        public static void OpenExistingAmendment(string documentPath, string templatePath, string documentAttachmentId, string templateAttachmentId, string documentName,
        string templateName, string versionId)
        {


            _strAmendmentDocumentPath = documentPath;
            _strNewAttachmentId = documentAttachmentId;
            _strAmendmentDocumentName = documentName;

            _strAmendmentTemplatePath = templatePath;
            _strAmendmentAttachmentId = templateAttachmentId;
            _strAmendmentTemplateName = templateName;
            _versionid = versionId;
            _d = Globals.ThisAddIn.getData();

            CompareSideBySide(_strAmendmentDocumentPath, _strAmendmentTemplatePath);

        }
        private static void UndoAllChanges(Word.Document doc)
        {
            object times = 1;
            int i = 0;
            while (doc.Undo(ref times))
            {
                Console.Write(i++);

            }
        }

        public static void TrackDocument()
        {
            try
            {


                objtempDocAmendment.TrackRevisions = false;

                object missing = System.Reflection.Missing.Value;
                //Microsoft.Office.Interop.Word.Range rngInsert = objtempAmendmentTemplate.Range(0, 0);
                //Microsoft.Office.Interop.Word.Range rngDelete = objtempAmendmentTemplate.Range(0, 0);
                //Microsoft.Office.Interop.Word.Range rngOther = objtempAmendmentTemplate.Range(0, 0);
                List<String> pText = new List<String>();

                //objtempAmendmentTemplate.RejectAllRevisions();
                UndoAllChanges(objtempAmendmentTemplate);

                // To get range \ location of axiom marker

                Word.Fields markerFields = objtempAmendmentTemplate.Fields;
                Word.Range rngField = objtempAmendmentTemplate.Range(0, 0);
                // bool isMarkerAvailable = false;
                bool firstMarkerFound = false;
                foreach (Microsoft.Office.Interop.Word.Field f in markerFields)
                {
                    if (f.Type == Word.WdFieldType.wdFieldMergeField)
                    {

                        if (!firstMarkerFound)
                        {
                            rngField = f.Code;
                            f.Delete();

                            firstMarkerFound = true;
                        }
                    }
                }
                Microsoft.Office.Interop.Word.Range rngInsert = objtempAmendmentTemplate.Range(rngField.Start, rngField.End);
                //end
                List<string> strIds = new List<string>();

                foreach (Microsoft.Office.Interop.Word.ContentControl contentControl in objtempDocAmendment.ContentControls)
                {
                    foreach (Word.Paragraph p in contentControl.Range.Paragraphs)
                    {
                        bool isTemplateTrackenabled = false;
                        if (objtempAmendmentTemplate.TrackRevisions == true)
                        {
                            objtempAmendmentTemplate.TrackRevisions = false;
                            isTemplateTrackenabled = true;
                        }

                        if (p.Range.Revisions.Count > 0)
                        {
                            contentControl.Range.Copy();
                            rngInsert.PasteSpecial();
                            if (isTemplateTrackenabled == true)
                            {
                                objtempAmendmentTemplate.TrackRevisions = true;
                            }
                            break;
                        }

                        if (isTemplateTrackenabled == true)
                        {
                            objtempAmendmentTemplate.TrackRevisions = true;
                        }
                    }

                }
                ///FIXME: Needs to change for paragraphs without clauses
                ///
                /// 
                    //foreach (Word.Paragraph p in objtempDocAmendment.Paragraphs)
                    //{
                    //    bool isTemplateTrackenabled = false;
                    //    if (objtempAmendmentTemplate.TrackRevisions == true)
                    //    {
                    //        objtempAmendmentTemplate.TrackRevisions = false;
                    //        isTemplateTrackenabled = true;
                    //    }

                    //    if (p.Range.ParentContentControl!=null) continue;

                    //    if (p.Range.Revisions.Count > 0)
                    //    {
                    //        p.Range.Copy();
                    //        rngInsert.PasteSpecial();
                    //    }

                    //    if (isTemplateTrackenabled == true)
                    //    {
                    //        objtempAmendmentTemplate.TrackRevisions = true;
                    //    }
                    //}
                ///
                ///


              

                //Remove Markup from template doc
                objtempAmendmentTemplate.TrackRevisions = false;
                Word.Fields fields = objtempAmendmentTemplate.Fields;
                foreach (Microsoft.Office.Interop.Word.Field f in fields)
                {
                    f.Select();
                    objtempAmendmentTemplate.Application.Selection.InsertParagraph();

                }
                objtempAmendmentTemplate.TrackRevisions = true;
                // End Remove Markup

             

            }
            catch (Exception exe)
            {
                //  MessageBox.Show(exe.ToString());

            }
            finally
            {
                if (objtempDocAmendment != null)
                {
                    objtempDocAmendment.TrackRevisions = true;
                    objtempDocAmendment.ActiveWindow.View.ShowRevisionsAndComments = false;
                }
            }
        }

     
        public DataReturn SaveCombinedDoc(string newAttachmentId, string fileAmendmentTemplatePath)
        {
            DataReturn dr = new DataReturn();
            try
            {
                string strFileAttached = fileAmendmentTemplatePath;
                //Save the Contract    

                //  DataReturn dr;
                _doc = Globals.ThisAddIn.Application.ActiveDocument;

                //Remove the contract refrence and id from document before save as combined doc.
                //To avoid save as a template
                Globals.ThisAddIn.AddDocId(_doc, "", "", "");
                //End

                _doc.SaveAs2(FileName: strFileAttached, FileFormat: Word.WdSaveFormat.wdFormatXMLDocument, CompatibilityMode: Word.WdCompatibilityMode.wdCurrent);
                string filenamecopy = AxiomIRISRibbon.Utility.SaveTempFile(_versionid + "Z");
                Word.Document dcopy = Globals.ThisAddIn.Application.Documents.Add(strFileAttached, Visible: false);
                dcopy.SaveAs2(FileName: filenamecopy, FileFormat: Word.WdSaveFormat.wdFormatXMLDocument, CompatibilityMode: Word.WdCompatibilityMode.wdCurrent);

                var docclose = (Word._Document)dcopy;
                docclose.Close();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(docclose);


                string vfilename = _versionName.Replace(" ", "_") + ".docx";
                dr = AxiomIRISRibbon.Utility.HandleData(_d.UpdateFile(newAttachmentId, vfilename, filenamecopy));

                return dr;
            }
            catch (Exception ex)
            {
                return dr;
            }
        }

        // Save Amend Document
        public static bool SaveAmend(bool ForceSave, bool SaveDoc, bool IsTemplate)
        {
            try
            {
                //System.Windows.Forms.MessageBox.Show("SaveAmend ");
                string strFileToSave, strVfilename, strAttachmentId;

                Globals.ThisAddIn.RemoveSaveHandler(); // remove the save handler to stop the save calling the save etc.

                Globals.ThisAddIn.ProcessingStart("Save Contract");
                DataReturn dr;
                _doc = Globals.ThisAddIn.Application.ActiveDocument;

                if (!IsTemplate)
                {
                    strFileToSave = _strAmendmentDocumentPath;
                    if (_versionName != null)
                    {
                        strVfilename = _versionName.Replace(" ", "_") + ".docx";
                    }
                    else { strVfilename = _strAmendmentDocumentName; }


                    strAttachmentId = _strNewAttachmentId;

                    // Globals.ThisAddIn.AddDocId(objtempDocAmendment, "Contract", "");
                }
                else
                {
                    strFileToSave = _strAmendmentTemplatePath;
                    if (_versionName != null)
                    {
                        strVfilename = _versionName + "_Amendment.docx";
                    }
                    else
                    {

                        strVfilename = _strAmendmentTemplateName;
                    }
                    strAttachmentId = _strAmendmentAttachmentId;

                    //   Globals.ThisAddIn.AddDocId(objtempAmendmentTemplate, "Contract", "");
                }




                if (SaveDoc)
                {

                    //save this to a scratch file
                    Globals.ThisAddIn.ProcessingUpdate("Save Scratch");
                    _doc.SaveAs2(FileName: strFileToSave, FileFormat: Word.WdSaveFormat.wdFormatXMLDocument, CompatibilityMode: Word.WdCompatibilityMode.wdCurrent);
                    //System.Windows.Forms.MessageBox.Show(" _doc.SaveAs2 ");

                    //Save a copy!
                    Globals.ThisAddIn.ProcessingUpdate("Save Copy");
                    string filenamecopy = AxiomIRISRibbon.Utility.SaveTempFile(_versionid + "Y");
                    Word.Document dcopy = Globals.ThisAddIn.Application.Documents.Add(strFileToSave, Visible: false);
                    dcopy.SaveAs2(FileName: filenamecopy, FileFormat: Word.WdSaveFormat.wdFormatXMLDocument, CompatibilityMode: Word.WdCompatibilityMode.wdCurrent);
                    //System.Windows.Forms.MessageBox.Show(" dcopy.SaveAs2 ");
                    var docclose = (Word._Document)dcopy;
                    docclose.Close();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(docclose);

                    Globals.ThisAddIn.ProcessingUpdate("Save To SalesForce");

                    dr = AxiomIRISRibbon.Utility.HandleData(_d.UpdateFile(strAttachmentId, strVfilename, filenamecopy));
                }
                Globals.ThisAddIn.AddSaveHandler(); // add it back in
                Globals.ThisAddIn.ProcessingStop("End");
                //System.Windows.Forms.MessageBox.Show("After End ");
                return true;
            }
            catch (Exception ex)
            {
                //System.Windows.Forms.MessageBox.Show("After End ");
                return true;
            }
        }


    }
}