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
        static Microsoft.Office.Interop.Word.Document objtempAmendmentTemplate;
        static Microsoft.Office.Interop.Word.Document objtempDocAmendment;

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
        private string _newVersionId;
        private string _strToAgreementId;
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
            try
            {

                if (this.radComboAmendment.SelectedItem == null && this.chkMaster.IsChecked == false)
                {
                    MessageBox.Show("Select any template from dropdown  or master checkbox");
                }
                else
                {
                    string strTemplateId = string.Empty;
                    if (chkMaster.IsChecked == true)
                    {
                        DataReturn dr = AxiomIRISRibbon.Utility.HandleData(_d.GetAmendmentTemplate(_strTemplate, true));
                        strTemplateId = dr.dt.Rows[0]["Id"].ToString();
                    }
                    else if (this.radComboAmendment.SelectedItem != null)
                    {
                        strTemplateId = ((RadComboBoxItem)(this.radComboAmendment.SelectedItem)).Tag.ToString();


                    }
                    if (strTemplateId != "selected")
                    {
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
                            CombineFiles(dtAllAttachments, _newVersionId, _strToAgreementId, _strTemplate, VersionName, VersionNumber, fileNameTemplate);
                            Globals.Ribbons.Ribbon1.CloseWindows();
                            this.Close();
                        }
                    }
                }
            }
            catch (Exception ex) { }
        }


        private void CombineFiles(DataTable dt, string versionid, string matterid, string templateid, string versionName, string versionNumber, string strFileNameTemplate)
        {
            try
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
                                // fileAmendmentDocument = _d.GetTempFilePath(rw["Id"].ToString() + "_" + rw["Name"].ToString());
                                fileAmendmentDocumentPath = _d.GetTempFilePath(rw["Id"].ToString() + _strSelectedAttachmentName);
                                File.WriteAllBytes(fileAmendmentDocumentPath, toBytes);
                                _strNewAttachmentId = rw["Id"].ToString();
                            }
                            else if (rw["Name"].ToString() == strFileNameTemplate)
                            {
                                byte[] toBytes = Convert.FromBase64String(rw["body"].ToString());
                                //  fileAmendmentTemplate = _d.GetTempFilePath(rw["Id"].ToString() + "_AmendmentTemplate");
                                fileAmendmentTemplatePath = _d.GetTempFilePath(rw["Id"].ToString() + "_" + rw["Name"].ToString());
                                File.WriteAllBytes(fileAmendmentTemplatePath, toBytes);
                                _fileToSaveAsAgreement = fileAmendmentTemplatePath;
                                _strAmendmentAttachmentId = rw["Id"].ToString();
                            }
                            // _source = app.Documents.Open(filename);
                        }
                    }
                    if (fileAmendmentDocumentPath == string.Empty && fileAmendmentTemplatePath == string.Empty)
                    {
                        MessageBox.Show("Files not avilable");
                    }
                    else
                    {
                        //  CombineDocs(fileAmendmentDocument, fileAmendmentTemplate, newAttachmentId);
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
                        DataReturn dr = SaveContract(_strNewAttachmentId, fileAmendmentTemplatePath);
                        app.Documents.Close();

                        //Open Files Sideby side
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
                        }
                        else if (rw["Name"].ToString() == fileNameTemplate)
                        {
                            byte[] toBytes = Convert.FromBase64String(rw["body"].ToString());
                            fileAmendmentTemplatePath = _d.GetTempFilePath(rw["Id"].ToString() + "_" + fileNameTemplate);
                            File.WriteAllBytes(fileAmendmentTemplatePath, toBytes);

                        }
                    }
                }
                // CompareSideBySide(fileAmendmentDocument, fileAmendmentTemplate);
                _strAmendmentTemplatePath = fileAmendmentTemplatePath;
                _strAmendmentDocumentPath = fileAmendmentDocumentPath;

                Microsoft.Office.Interop.Word.Application app = Globals.ThisAddIn.Application;

                object newFilenameObject1 = fileAmendmentTemplatePath;
                objtempAmendmentTemplate = app.Documents.Open(ref newFilenameObject1, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing);

                object newFilenameObject2 = fileAmendmentDocumentPath;
                objtempDocAmendment = app.Documents.Open(ref newFilenameObject2, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);

                //AmendmentTemplate
                Globals.ThisAddIn.AddDocId(objtempAmendmentTemplate, "AmendmentTemplate", "");
                //AmendmentDocument
                Globals.ThisAddIn.AddDocId(objtempDocAmendment, "AmendmentDocument", "");

                object o = objtempAmendmentTemplate;
                objtempDocAmendment.Windows.CompareSideBySideWith(ref o);
                objtempDocAmendment.AcceptAllRevisions();
                objtempDocAmendment.TrackRevisions = true;
                objtempAmendmentTemplate.TrackRevisions = true;
            }
        }

        public static void TrackDocument()
        {
            try
            {
                Object destFile = _strAmendmentTemplatePath;
                int delcnt = 0;
                int inscnt = 0;
                string wrInsert = string.Empty;
                string wrOthers = string.Empty;
                Microsoft.Office.Interop.Word.Range rngInsert = objtempAmendmentTemplate.Range(0, 0);
                Microsoft.Office.Interop.Word.Range rngDelete = objtempAmendmentTemplate.Range(0, 0);
                objtempAmendmentTemplate.RejectAllRevisions();

                foreach (Microsoft.Office.Interop.Word.Section s in objtempDocAmendment.Sections)
                {
                    for (int rnumber = 1; rnumber <= s.Range.Revisions.Count; rnumber++)
                    {
                        Microsoft.Office.Interop.Word.Revision r = s.Range.Revisions[rnumber];

                        if (r.Type == Microsoft.Office.Interop.Word.WdRevisionType.wdRevisionInsert) // Inserted
                        {
                            rngInsert.Font.StrikeThrough = 0;
                            inscnt += r.Range.Words.Count;
                            wrInsert = r.Range.Text;
                            rngInsert.Text += wrInsert;
                            rngInsert.Text += "\u000A";
                            rngInsert.Bold = 1;
                            rngInsert.Font.StrikeThrough = 0;
                        }
                        if (r.Type == Microsoft.Office.Interop.Word.WdRevisionType.wdRevisionDelete) // Deleted
                        {
                            delcnt += r.Range.Words.Count;
                            wrOthers = r.Range.Text;
                            //Microsoft.Office.Interop.Word.Range rng = objtempAmendmentTemplate.Range(0, 0);
                            rngDelete.Font.StrikeThrough = 0;
                            rngDelete.Text += wrOthers;
                            rngDelete.Text += "\u000A";
                            rngDelete.Font.StrikeThrough = 1;
                          //  CopyDeletedContent(objtempAmendmentTemplate, wrOthers);
                        }

                        /*  if (r.Type == Microsoft.Office.Interop.Word.WdRevisionType.wdRevisionProperty) // Formatting (bold,italics)
                          {
                              inscnt += r.Range.Words.Count;
                             // Microsoft.Office.Interop.Word.Range rng = objtempAmendmentTemplate.Range(0, 0);
                              rngInsert.Text = wr;
                          }*/
                    }
                }
            }
            catch (Exception ex) { }
        }
        private static void CopyDeletedContent(Microsoft.Office.Interop.Word.Document doc, string deletedcontent)
        {
            foreach (Microsoft.Office.Interop.Word.Field myMergeField in doc.Fields)
            {
                Microsoft.Office.Interop.Word.Range rngFieldCode = myMergeField.Code;

                String fieldText = rngFieldCode.Text.Trim();


                if (fieldText.StartsWith("MERGEFIELD  Name  \\* MERGEFORMAT"))
                {
                    Int32 endMerge = fieldText.IndexOf("\\");

                    Int32 fieldNameLength = fieldText.Length - endMerge;

                    String fieldName = fieldText.Substring(11, endMerge - 11);


                    fieldName = fieldName.Trim();

                    // **** FIELD REPLACEMENT IMPLEMENTATION GOES HERE ****//

                    // THE PROGRAMMER CAN HAVE HIS OWN IMPLEMENTATIONS HERE

                    if (fieldName == "Name")
                    {
                        myMergeField.Select();
                        //myMergeField.Application.Selection.InsertBreak();
                        //Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
                        myMergeField.Application.Selection.InsertAfter("\n");

                        myMergeField.Application.Selection.InsertAfter(deletedcontent);
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
        public static bool SaveContract(bool ForceSave, bool SaveDoc, bool IsTemplate)
        {
            string strFileToSave,strVfilename,strAttachmentId;
            if (!IsTemplate)
            {
                strFileToSave = _strAmendmentDocumentPath;
                strVfilename = _versionName.Replace(" ", "_") + ".docx";
                strAttachmentId = _strNewAttachmentId;
            }
            else
            {
                strFileToSave = _strAmendmentTemplatePath;
                strVfilename = _versionName + "_Amendment.docx";
                strAttachmentId = _strAmendmentAttachmentId;
            }

            
            Globals.ThisAddIn.RemoveSaveHandler(); // remove the save handler to stop the save calling the save etc.

            Globals.ThisAddIn.ProcessingStart("Save Contract");
            DataReturn dr;
            _doc = Globals.ThisAddIn.Application.ActiveDocument;

            if (SaveDoc)
            {

                //save this to a scratch file
                Globals.ThisAddIn.ProcessingUpdate("Save Scratch");
                _doc.SaveAs2(FileName: strFileToSave, FileFormat: Word.WdSaveFormat.wdFormatXMLDocument, CompatibilityMode: Word.WdCompatibilityMode.wdCurrent);

                //Save a copy!
                Globals.ThisAddIn.ProcessingUpdate("Save Copy");
                string filenamecopy = AxiomIRISRibbon.Utility.SaveTempFile(_versionid + "X");
                Word.Document dcopy = Globals.ThisAddIn.Application.Documents.Add(strFileToSave, Visible: false);
                dcopy.SaveAs2(FileName: filenamecopy, FileFormat: Word.WdSaveFormat.wdFormatXMLDocument, CompatibilityMode: Word.WdCompatibilityMode.wdCurrent);

                var docclose = (Microsoft.Office.Interop.Word._Document)dcopy;
                docclose.Close();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(docclose);

                Globals.ThisAddIn.ProcessingUpdate("Save To SalesForce");

                dr = AxiomIRISRibbon.Utility.HandleData(_d.UpdateFile(strAttachmentId, strVfilename, filenamecopy));

            }
            Globals.ThisAddIn.AddSaveHandler(); // add it back in
            Globals.ThisAddIn.ProcessingStop("End");
            return true;
        }


    }
}