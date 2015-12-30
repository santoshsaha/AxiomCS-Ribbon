﻿﻿using System;
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

        void CreateAndLoadAmendment_AsyncRun(object sender, DoWorkEventArgs args, string amdTplId)
        {
            try
            {
                _versionNumber = _versionNumber + 1;
                _versionName = "Version " + _versionNumber;

                // Create amendment version record
                DataReturn created = AxiomIRISRibbon.Utility.HandleData(
                                     _d.CreateVersion(String.Empty, _strToAgreementId, _strTemplate, _versionName, Convert.ToString(_versionNumber), _allDr));
                _newVersionId = created.id;

                // Bring over all attachments in last version
                DataReturn dr = AxiomIRISRibbon.Utility.HandleData(_d.GetVersionAllAttachments(_versionid));
                if (!dr.success)
                {
                    MessageBox.Show("Error [AMND003] creating amendment; Failed to retrieve attachments");
                    args.Cancel = true;
                    return;
                }
                DataTable attachments = dr.dt;
                if (attachments.Rows.Count == 0)
                {
                    MessageBox.Show("Error [AMND004] creating amendment; No attachments found for version " + (_versionNumber - 1));
                    args.Cancel = true;
                    return;
                }

                // Look for the agreement document attachment
                string localAgrFilename = null;
                foreach (DataRow rw in attachments.Rows)
                {
                    string filename = rw["Name"].ToString();
                    // FIXME: Need better scheme to identify agreement document
                    if (filename == "Version_" + (_versionNumber - 1) + ".docx")
                    {
                        _strSelectedAttachmentName = _versionName.Replace(" ", "_") + ".docx";
                        string body = rw["body"].ToString();
                        // Save agreement document to SF
                        DataReturn dr2 = _d.saveAttachmentstoSF(_newVersionId, _strSelectedAttachmentName, body);
                        if (!dr2.success)
                        {
                            MessageBox.Show("Error [AMND013] creating amendment; Failure saving agreement: " + dr2.errormessage);
                            args.Cancel = true;
                            return;
                        }
                        _strNewAttachmentId = dr2.id;
                        // Identify and set temp file name for temporary document
                        string tmpPath = System.IO.Path.GetTempPath();
                        localAgrFilename = tmpPath + "\\" + _versionName.Replace(" ", "_") + ".docx";
                        if (System.IO.File.Exists(localAgrFilename))
                        {
                            try
                            {
                                System.IO.File.Delete(localAgrFilename);
                            }
                            catch (UnauthorizedAccessException uae)
                            {
                                MessageBox.Show("Error [AMND006] creating amendment; Permission denied for temp directory; " + uae.Message);
                                args.Cancel = true;
                                return;
                            }
                        }
                        _strAmendmentDocumentPath = localAgrFilename;
                        _strAmendmentDocumentName = _versionName.Replace(" ", "_") + ".docx";
                        // Save the agreement document to temp location
                        byte[] agrBytes = Convert.FromBase64String(body);
                        File.WriteAllBytes(localAgrFilename, agrBytes);
                        break;
                    }
                }
                if (localAgrFilename == null)
                {
                    MessageBox.Show("Error [AMND005] creating amendment; No agreement document found for version " + (_versionNumber - 1));
                    args.Cancel = true;
                    return;
                }
                // Retrieve and save amendment template as amendment document into SF
                dr = AxiomIRISRibbon.Utility.HandleData(_d.GetTemplateAttach(amdTplId));
                if (!dr.success)
                {
                    MessageBox.Show("Error [AMND007] creating amendment; Failed to retrieve amendment template");
                    args.Cancel = true;
                    return;
                }
                attachments = dr.dt;
                if (attachments.Rows.Count == 0)
                {
                    MessageBox.Show("Error [AMND004] creating amendment; No attachments found for amendment template ");
                    args.Cancel = true;
                    return;
                }
                string tplFilename = _versionName.Replace(" ", "_") + "_Amendment.docx";
                // FIXME: Assume the tpl doc is the first attachment
                string amdBody = attachments.Rows[0]["body"].ToString();
                DataReturn dr3 = _d.saveAttachmentstoSF(_newVersionId, tplFilename, amdBody);
                if (!dr3.success)
                {
                    MessageBox.Show("Error [AMND013] creating amendment; Failure saving amendment: " + dr3.errormessage);
                    args.Cancel = true;
                    return;
                }
                _strAmendmentAttachmentId = dr3.id;

                // Save amendment document to temp location
                string localTplFilename = System.IO.Path.GetTempPath() + "\\" + tplFilename;
                if (System.IO.File.Exists(localTplFilename))
                {
                    try
                    {
                        System.IO.File.Delete(localTplFilename);
                    }
                    catch (UnauthorizedAccessException uae)
                    {
                        MessageBox.Show("Error [AMND008] creating amendment; Permission denied for temp directory; " + uae.Message);
                        args.Cancel = true;
                        return;
                    }
                }
                _strAmendmentTemplatePath = localTplFilename;
                _strAmendmentTemplateName = tplFilename;
                // Save the amendment document to temporary location
                byte[] amdBytes = Convert.FromBase64String(amdBody);
                File.WriteAllBytes(localTplFilename, amdBytes);

                // Add the two doc names to args so they can be accessed in the complete callback
                args.Result = new CompareAmendmentArgs(localAgrFilename, localTplFilename);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error [AMND002] creating amendment; " + ex.Message);
                args.Cancel = true;
            }
        }
        // AH
        void CreateAndLoadAmendment_AsyncCompleted(object sender, RunWorkerCompletedEventArgs args)
        {
            try
            {
                if (args.Cancelled) return;

                if (args.Result == null)
                {
                    MessageBox.Show("Error [AMND010] showing amendment; No result");
                    return;
                }
                object agreementFileName = ((CompareAmendmentArgs)args.Result).AgreementFileName;
                object amendmentFileName = ((CompareAmendmentArgs)args.Result).AmendmentFileName;

                // Close current active document
                ((Microsoft.Office.Interop.Word._Document)Globals.ThisAddIn.Application.ActiveDocument).Close();

                // Open documents
                object missing = System.Reflection.Missing.Value;
                Word.Documents documents = Globals.ThisAddIn.Application.Documents;
                Word.Document agreement = documents.Open(ref agreementFileName, ref missing, ref missing, ref missing, ref missing,
                                                         ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                                                         ref missing, ref missing, ref missing, ref missing, ref missing);
                Word.Document amendment = documents.Open(ref amendmentFileName, ref missing, ref missing, ref missing, ref missing,
                                                         ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                                                         ref missing, ref missing, ref missing, ref missing, ref missing);

                //To hide Marker. Uncomment for prod move
                /*Word.Fields amdFields = amendment.Fields;
                Word.Range insMarker = null;
                foreach (Word.Field f in amdFields)
                {
                    if (f.Type == Word.WdFieldType.wdFieldMergeField && f.Result != null && f.Result.Text == "«AxiomMarker»")
                    {
                        insMarker = f.Result;

                        //To hide axiom marker
                        insMarker.Font.ColorIndex = Word.WdColorIndex.wdWhite;
                        amendment.ActiveWindow.View.FieldShading = 0;
                        break;
                    }
                }
                */
                // Unlock agreement for edits
                for (int i = 1; i <= agreement.ContentControls.Count; i++)
                {
                    agreement.ContentControls[i].LockContents = false;
                    agreement.ContentControls[i].LockContentControl = false;
                }

                // Unlock for edits
                for (int i = 1; i <= amendment.ContentControls.Count; i++)
                {
                    amendment.ContentControls[i].LockContents = false;
                    amendment.ContentControls[i].LockContentControl = false;
                }

                // Add property for saves
                Globals.ThisAddIn.AddDocId(amendment, "Contract", "", "AmendmentTemplate");
                Globals.ThisAddIn.AddDocId(agreement, "Contract", "", "AmendmentDocument");

                // Show compare side-by-side
                agreement.TrackRevisions = true;
                agreement.ActiveWindow.View.ShowRevisionsAndComments = true;
                amendment.TrackRevisions = true;

                amendment.Activate();
                object o = agreement;
                amendment.Windows.CompareSideBySideWith(ref o);
                // Reposition side by side
                // FIXME: Side by side documents interchanged. Need to fix
                amendment.Windows.ResetPositionsSideBySide();

               agreement.ActiveWindow.View.Type = Word.WdViewType.wdPrintView;

                Globals.Ribbons.Ribbon1.CloseWindows();
                this.windowAttachmentsView.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error [AMND009] opening full/amendment views; " + ex.Message);
            }
        }
        // AH





        private class CompareAmendmentArgs
        {

            public string AgreementFileName { get; private set; }
            public string AmendmentFileName { get; private set; }

            public CompareAmendmentArgs(string agrFN, string amdFN)
            {
                AgreementFileName = agrFN;
                AmendmentFileName = amdFN;
            }

        }
        // AH
        private void btnOpen_Click(object sender, RoutedEventArgs args)
        {
            try
            {
                // Read selected amendment template
                if (this.radComboAmendment.SelectedItem == null && this.chkMaster.IsChecked == false)
                {
                    MessageBox.Show("Please select the amendment template from the dropdown or select master checkbox");
                    return;
                }
                string amdTplId = null;
                if (chkMaster.IsChecked == true)
                {
                    DataReturn dr = AxiomIRISRibbon.Utility.HandleData(_d.GetAmendmentTemplate(null, true));
                    amdTplId = dr.dt.Rows[0]["Id"].ToString();
                }
                else if (this.radComboAmendment.SelectedItem != null)
                {
                    amdTplId = ((RadComboBoxItem)(this.radComboAmendment.SelectedItem)).Tag.ToString();
                    if (amdTplId == "select")
                    {
                        MessageBox.Show("Please select the amendment template from the dropdown or select master checkbox");
                        return;
                    }
                }
                // Set busy indicator
                bsyCompareAmndIndc.IsBusy = true;
                bsyCompareAmndIndc.BusyContent = "Loading ...";
                // Create worker and run
                bsyCompareAmndIndicatorBackgroundWorker = new BackgroundWorker();
                bsyCompareAmndIndicatorBackgroundWorker.DoWork += (obj, ev) => CreateAndLoadAmendment_AsyncRun(obj, ev, amdTplId);
                bsyCompareAmndIndicatorBackgroundWorker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(CreateAndLoadAmendment_AsyncCompleted);
                bsyCompareAmndIndicatorBackgroundWorker.RunWorkerAsync();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error [AMND001] creating amendment; " + ex.Message);
            }
        }


        public static void OpenExistingAmendment(string documentPath, string templatePath, string documentAttachmentId, string templateAttachmentId, string documentName,
        string templateName, string versionId)
        {
            MessageBox.Show("This operation is not yet supported");

            _strAmendmentDocumentPath = documentPath;
            _strNewAttachmentId = documentAttachmentId;
            _strAmendmentDocumentName = documentName;

            _strAmendmentTemplatePath = templatePath;
            _strAmendmentAttachmentId = templateAttachmentId;
            _strAmendmentTemplateName = templateName;
            _versionid = versionId;
            _d = Globals.ThisAddIn.getData();

            //CompareSideBySide(_strAmendmentDocumentPath, _strAmendmentTemplatePath);

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

        // FIXME: Refactor to controller class
        public static void TrackDocument()
        {
            Word.Document agreement = null;
            Word.Document amendment = null;
            try
            {
                Word.Documents documents = Globals.ThisAddIn.Application.Documents;

                DateTime lastAmendedDate = DateTime.MinValue;
                // FIXME: Need better scheme to discover agreement and amendment
                foreach (Word.Document d in documents)
                {
                    if (amendment == null && d.FullName.Contains("_Amendment")) amendment = d;
                    else if (agreement == null && d.FullName.Contains("Version")) agreement = d;
                    if (agreement != null && amendment != null) break;
                }
                string ts = string.Empty;
                if (amendment != null)
                {
                    ts = Globals.ThisAddIn.GetDocTimeStamp(amendment);
                }
                // It is critical to turn off tracking when moving changes over for
                // markup to show correctly in the amendment document
                agreement.TrackRevisions = false;
                amendment.TrackRevisions = false;
                if (!string.IsNullOrEmpty(ts))
                {
                    lastAmendedDate = DateTime.ParseExact(ts, "yyyy-MM-dd HH:mm:ss.fff tt", null);
                }
                Globals.ThisAddIn.SetDocTimeStamp(amendment);

                HashSet<string> seen = new HashSet<string>();
                foreach (Word.Revision r in agreement.Revisions)
                {
                    if (lastAmendedDate > r.Date) continue;

                    Word.Range insPosition = GetAmendmentDocumentInsertPosition(amendment);
                    if (insPosition == null)
                    {
                        MessageBox.Show("Error [AMND012] while syncing; No insert marker found in amendment document");
                        return;
                    }

                    // FIXME: Handle deleted revision
                    if (r.Type == Word.WdRevisionType.wdRevisionDelete)
                    {
                    }

                    // If whitespace, insert only that space, ignoring the paragraph and surrounding text
                    if (r.Range.Text.Trim() == String.Empty)
                    {
                        insPosition.InsertBefore(r.Range.Text);
                    }
                    // If range has content controls, handle each
                    else if (r.Range.ContentControls !=null && r.Range.ContentControls.Count > 0)
                    {
                        foreach (Word.ContentControl cc in r.Range.ContentControls)
                        {
                            //FIXME: Need better scheme to determine Header and Signature
                            if (cc.Title.StartsWith("Header") || cc.Title.StartsWith("Signature")) continue;

                            // If CC is a child content control, skip - handle move via parent
                            if (cc.ParentContentControl != null) continue;

                            // Skip if already seen
                            if (seen.Contains(cc.ID)) continue;
                            seen.Add(cc.ID);

                            cc.Copy();
                            insPosition.Collapse();
                            insPosition.Paste();

                            // FIXME: Optimize - replace repetative detection of marker
                            insPosition = GetAmendmentDocumentInsertPosition(amendment);
                            if (insPosition == null)
                            {
                                MessageBox.Show("Error [AMND012] while syncing; No insert marker found in amendment document");
                                return;
                            }
                        }
                    }
                    // If range is inside a content control, bring over the entire content control
                    else if (r.Range.ParentContentControl != null)
                    {
                        Word.ContentControl cc = r.Range.ParentContentControl;

                        //FIXME: Need better scheme to determine Header and Signature
                        if (cc.Title.StartsWith("Header") || cc.Title.StartsWith("Signature")) continue;

                        // Skip if already seen
                        if (seen.Contains(cc.ID)) continue;
                        seen.Add(cc.ID);

                        cc.Copy();
                        insPosition.Collapse(); // Required to prevent Command Failed error on Paste under certain circumstances
                        insPosition.Paste();
                    }
                    else if (r.Range.Paragraphs.Count > 0)
                    {
                        foreach (Word.Paragraph p in r.Range.Paragraphs)
                        {
                            p.Range.Copy();
                            insPosition.Collapse();
                            insPosition.Paste();
                            insPosition = GetAmendmentDocumentInsertPosition(amendment);
                            if (insPosition == null)
                            {
                                MessageBox.Show("Error [AMND012] while syncing; No insert marker found in amendment document");
                                return;
                            }
                        }
                    }
                    else
                    {
                        MessageBox.Show("Error [AMDN015] while syncing; Range has no context: " + r.Range.Text);
                    }
                }

                // Move insert marker
                amendment.TrackRevisions = false;
                Word.Fields amdMergeFields = amendment.Fields;
                Word.Range insAmendMarker = null;
                //Word.Range insPosition = null;
                foreach (Word.Field f in amdMergeFields)
                {
                    if (f.Type == Word.WdFieldType.wdFieldMergeField && f.Result != null && f.Result.Text == "«AxiomMarker»")
                    {
                        insAmendMarker = f.Result;
                        f.Code.InsertBefore(Environment.NewLine);
                        break;
                    }
                }
                amendment.TrackRevisions = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error [AMND011] syncing amendment; " + ex.Message);
            }
            finally
            {
                if (agreement != null) agreement.TrackRevisions = true;
                if (amendment != null) amendment.TrackRevisions = true;
            }
        }


        private static Word.Range GetAmendmentDocumentInsertPosition(Word.Document amendment)
        {
            // Locate insert marker in amendment document
            Word.Fields amdFields = amendment.Fields;
            Word.Range insMarker = null;
            Word.Range insPosition = null;
            foreach (Word.Field f in amdFields)
            {
                if (f.Type == Word.WdFieldType.wdFieldMergeField && f.Result != null && f.Result.Text == "«AxiomMarker»")
                {
                    insMarker = f.Result;
                    f.Code.InsertBefore(Environment.NewLine);
                    object wdth = Environment.NewLine.Length;
                    insPosition = f.Code.Previous(Word.WdUnits.wdCharacter, ref wdth);
                    return insPosition;
                }
            }
            return null;
        }

        // AH
        // FIXME: Refactor to controller class
        /*
        public static void TrackDocument()
        {
            try
            {
                Word.Documents documents = Globals.ThisAddIn.Application.Documents;
                Word.Document agreement = null;
                Word.Document amendment = null;

                // FIXME: Need better scheme to discover agreement and amendment
                foreach (Word.Document d in documents)
                {
                    if (amendment == null && d.FullName.Contains("_Amendment")) amendment = d;
                    else if (agreement == null && d.FullName.Contains("Version")) agreement = d;
                    if (agreement != null && amendment != null) break;
                }


                // Look for mods in agreement document
                foreach (Word.ContentControl clause in agreement.ContentControls)
                {
                    // FIXME: Need better scheme to determine header and signature
                    if (clause.Title.StartsWith("Header") || clause.Title.StartsWith("Signature")) continue;

                    // Ignore unchanged clauses
                    if (clause.Range.Revisions.Count == 0) continue;

                    // FIXME: Handle deletion of clause

                    // FIXME: Optimize - replace repetative detection of marker
                    // Locate insert marker in amendment document
                    Word.Fields amdFields = amendment.Fields;
                    Word.Range insMarker = null;
                    foreach (Word.Field f in amdFields)
                    {
                        if (f.Type == Word.WdFieldType.wdFieldMergeField && f.Result != null && f.Result.Text == "«AxiomMarker»")
                        {
                            insMarker = f.Result;
                            break;
                        }
                    }
                    if (insMarker == null)
                    {
                        MessageBox.Show("Error [AMND012] while syncing; No insert marker found in amendment document");
                        return;
                    }
                    // Make space
                    insMarker.InsertBefore("\r\n");
                    object wdth = "\r\n".Length;
                    Word.Range insPosition = insMarker.Previous(Word.WdUnits.wdCharacter, ref wdth);
                    // Copy the content control over
            
                    //object r = insPosition;
                    //Word.ContentControl amdClause = amendment.ContentControls.Add(clause.Type, ref r);
                    //amdClause.BuildingBlockCategory = clause.BuildingBlockCategory;
                    //amdClause.BuildingBlockType = clause.BuildingBlockType;
                    //amdClause.Checked = clause.Checked;
                    //amdClause.DateCalendarType = clause.DateCalendarType;
                    //amdClause.DateDisplayFormat = clause.DateDisplayFormat;
                    //amdClause.DateDisplayLocale = clause.DateDisplayLocale;
                    //amdClause.DateStorageFormat = clause.DateStorageFormat;

                    //amdClause.Title = clause.Title;
                    //amdClause.Tag = clause.Tag;
                    //amdClause.Range.Text = clause.Range.Text;
                    
                    clause.Copy();
                    insPosition.Paste();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error [AMND011] syncing amendment; " + ex.Message);
            }
        }
       */     

        public static void TrackDocumentOld()
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

                //Globals.ThisAddIn.ProcessingStart("Save Contract");
                DataReturn dr;

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
                        strVfilename = _versionName.Replace(" ", "_") + "_Amendment.docx";
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
                    _doc = Globals.ThisAddIn.Application.Documents[strVfilename];
                    //save this to a scratch file
                    //Globals.ThisAddIn.ProcessingUpdate("Save Scratch");
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