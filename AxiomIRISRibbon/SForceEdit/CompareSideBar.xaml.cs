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
using System.IO;


namespace AxiomIRISRibbon.SForceEdit
{
    /// <summary>
    /// Interaction logic for SForceEditSideBar.xaml
    /// New File added by PES
    /// </summary>
    public partial class CompareSideBar : UserControl
    {

        private static Data _d;
        static Word.Application app;
        static Word.Document _source;
        private static string _fileName;
        private static string _matterid;
        private static string _versionid;
        private static string _templateid;
        private static string _versionName;
        private static string _versionNumber;
        private static string _attachmentid;
        private static Word.Document _doc;
        private static bool _firstsave;

        public CompareSideBar()
        {
            InitializeComponent();
            AxiomIRISRibbon.Utility.setTheme(this);

            _d = Globals.ThisAddIn.getData();

            LoadTemplatesDLL();

        }      

        public void Create(string filename, string versionid, string matterid, string templateid, string versionName, string versionNumber, string attachmentid)
        {
            _fileName = filename;
            _matterid = matterid;
            _versionid = versionid;
            _templateid = templateid;
            _versionName = versionName;
            _versionNumber = versionNumber;
            _attachmentid = attachmentid;

        }

        private void LoadTemplatesDLL()
        {
            try
            {

                DataReturn dr = AxiomIRISRibbon.Utility.HandleData(_d.GetAgreementTemplates(true));

                if (!dr.success) return;

                DataTable dt = dr.dt;
                cbTemplates.Items.Clear();

                RadComboBoxItem i;

                foreach (DataRow r in dt.Rows)
                {
                    i = new RadComboBoxItem();
                    i.Tag = r["Id"].ToString();
                    i.Content = r["Name"].ToString();
                    this.cbTemplates.Items.Add(i);

                }

            }
            catch (Exception ex)
            {
            }
        }

        private void btnCompare_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                btnCompare.IsEnabled = false;
                object missing = System.Reflection.Missing.Value;

                if (this.cbTemplates.SelectedItem != null)
                {

                    string TemplateId = ((RadComboBoxItem)(this.cbTemplates.SelectedItem)).Tag.ToString();
                    string TemplateName = ((RadComboBoxItem)(this.cbTemplates.SelectedItem)).Content.ToString();

                    DataReturn drAttachemnts = AxiomIRISRibbon.Utility.HandleData(_d.GetTemplateAttach(TemplateId));
                    if (!drAttachemnts.success) return;

                    DataTable dtAttachments = drAttachemnts.dt;
                    string fileTemplate = "";
                    foreach (DataRow rw in dtAttachments.Rows)
                    {
                        byte[] toBytes = Convert.FromBase64String(rw["body"].ToString());
                        fileTemplate = _d.GetTempFilePath("Template View");
                        File.WriteAllBytes(fileTemplate, toBytes);
                    }

                    Word.Document wordTemplate;
                    Word.Document wordAttachment;
                    Word.Application app = Globals.ThisAddIn.Application;

                    wordAttachment = app.Documents[app.ActiveDocument.FullName]; // Document already open

                    object objTemplate = fileTemplate;
                    // FIXME: Template was not loading in right side, if file size is more. So added thread.sleep
                    System.Threading.Thread.Sleep(8000);

                    wordTemplate = app.Documents.Open(ref objTemplate, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);
  
                    // Unlock clauses in template - track changes in agreement requires unlocked clauses
                    // Right side is RO in Document.Compare, so template still cannot be modified
                    for (int i = 1; i <= wordTemplate.ContentControls.Count; i++)
                    {
                        wordTemplate.ContentControls[i].LockContents = false;
                        wordTemplate.ContentControls[i].LockContentControl = false;
                    }

                    // Code to remove   the document modified by -- This code will reset all the properties in the document
                    //wordTemplate.RemoveDocumentInformation(Microsoft.Office.Interop.Word.WdRemoveDocInfoType.wdRDIDocumentProperties);

                    // End Code

                    //Compare
                    Globals.ThisAddIn.AddDocId(wordTemplate, "Contract", "", "Compare");
                    
                    //added below lines to close the open file before opening the split screen
                    // First collect files to close and then close; closing files while iterating 
                    // through app.Documents enumerable removes from collection and files are missed
                    List<Word._Document> docsToClose = new List<Word._Document>();
                    foreach (Word._Document d in app.Documents)
                    {
                        if (d.FullName != fileTemplate) docsToClose.Add(d);
                    }
                    object saveOption = Word.WdSaveOptions.wdDoNotSaveChanges;
                    object originalFormat = Word.WdOriginalFormat.wdOriginalDocumentFormat;
                    object routeDocument = false;
                    foreach (Word._Document d in docsToClose) d.Close(ref saveOption, ref originalFormat, ref routeDocument);

                    //  Compare code
                    wordTemplate.Compare(_fileName, missing, Word.WdCompareTarget.wdCompareTargetNew, true, true, false, true, false);
                    // Unlock agreement clauses for edit; Must be done on the active window document
                    // Some clauses (EventsOfDefault) do not unlock when unlocking directly using the Document reference
                    foreach (Word.ContentControl cc in app.ActiveDocument.ContentControls)
                    {
                        cc.LockContents = false;
                        cc.LockContentControl = false;
                    }
                    app.ActiveWindow.View.SplitSpecial = Word.WdSpecialPane.wdPaneRevisionsVert;
                    app.ActiveWindow.ShowSourceDocuments = Word.WdShowSourceDocuments.wdShowSourceDocumentsOriginal;
                    app.ActiveDocument.TrackRevisions = true;
                    app.ActiveDocument.TrackMoves = true;
                    app.ActiveDocument.TrackFormatting = true;
                    // FIXME: Perhaps not required if track changes turned off when generating version 1 agreement doc (New from Existing)
                    app.ActiveDocument.AcceptAllRevisionsShown();

                    //Code to resize review panel
                    app.ActiveWindow.Document.Frameset.ChildFramesetItem[2].WidthType = Word.WdFramesetSizeType.wdFramesetSizeTypePercent;
                    app.ActiveWindow.Document.Frameset.ChildFramesetItem[2].Width = 82;  // Standard size 75. + will reduce size and - will increase size

                    app.Activate();
                    // Russel Dec11 - add in the Doc Id to the comparison doc
                    Globals.ThisAddIn.AddDocId(app.ActiveDocument, "Contract", "", "Compare");

                    // close the temp files
                    var docTemplateClose = (Word._Document)wordTemplate;
                    docTemplateClose.Close(SaveChanges: false);

                    //End Compare
                    Globals.Ribbons.Ribbon1.CloseWindows();

                }
                else
                {
                    MessageBox.Show("Please select a template");

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error generating Compare");

            }
            finally
            {
                btnCompare.IsEnabled = true;
            }
        }

  
        public static bool SaveCompare(bool ForceSave, bool SaveDoc)
        {
            try
            {
                //  int seq = 1;
                string strFileAttached = _fileName;
                if (strFileAttached != null)
                {
                    //Save the Contract    
                    Globals.ThisAddIn.RemoveSaveHandler(); // remove the save handler to stop the save calling the save etc.

                    // if this is the first save - then save everything
                    if (_firstsave) ForceSave = true;

                    Globals.ThisAddIn.ProcessingStart("Save Contract");
                    DataReturn dr;
                    _doc = Globals.ThisAddIn.Application.ActiveDocument;

                    /*
                    Globals.ThisAddIn.ProcessingUpdate("Get the Clause Values");
                    foreach (Word.ContentControl cc in _doc.Range().ContentControls)
                    {
                        if (cc.Tag != null)
                        {
                            string tag = cc.Tag;
                            string docclausename = cc.Title;
                            Globals.ThisAddIn.ProcessingUpdate("Save " + docclausename);

                            string text = cc.Range.Text;
                            string[] taga = cc.Tag.Split('|');
                            if (taga[0] == "Concept")
                            {
                                dr = AxiomIRISRibbon.Utility.HandleData(_d.SaveDocumentClause("", _versionid, taga[1], taga[2], docclausename, seq++, text, false));
                            }
                            else if (taga[0] == "element")
                            {

                                dr = AxiomIRISRibbon.Utility.HandleData(_d.SaveDocumentClauseElement("",docclausename ,taga[2],"",_templateid,text,text ));
                                                      
                            }
                        }
                    }

                    */

                    /*   List<string> clauseArray = new List<string>();
                       DataReturn ribbonClauseDr = _d.getRibbonClause(_templateid);
                       //check condition
                       DataTable ribbonClauseDt = ribbonClauseDr.dt;
                       foreach (Word.ContentControl cc in _doc.Range().ContentControls)
                       {
                           //
                           if (cc.Tag != null)
                           {
                               //
                               string tag = cc.Tag;
                               string docclausename = cc.Title;
                           }
                           string text = cc.Range.Text;
                           string[] taga = cc.Tag.Split('|');
                           if (taga[0] == "Concept")
                           {
                               //
                               if (!clauseArray.Contains(taga[2]))
                               {
                                   //
                                   clauseArray.Add(taga[2]);
                               }
                               //dr = _d.SaveRibbonClause(ribbonClauseDt, taga[1], taga[2], _templateid);
                           }
                       }
                       if (clauseArray.Count > 0)
                       {
                           dr = _d.SaveRibbonClause(ribbonClauseDt, clauseArray, _templateid);
                       }

                       */



                    dr = AxiomIRISRibbon.Utility.HandleData(_d.SaveVersion(_versionid, _matterid, _templateid, _versionName, _versionNumber));
                    if (!dr.success) return false;
                    _versionid = dr.id;

                    if (SaveDoc)
                    {
                        //Save the file as an attachment
                        //save this to a scratch file
                        Globals.ThisAddIn.ProcessingUpdate("Save Scratch");
                        _doc.SaveAs2(FileName: strFileAttached, FileFormat: Word.WdSaveFormat.wdFormatXMLDocument, CompatibilityMode: Word.WdCompatibilityMode.wdCurrent);

                        //Save a copy because document open in split view cannot be locked for upload
                        //into IRIS
                        Globals.ThisAddIn.ProcessingUpdate("Save Copy");
                        string filenamecopy = AxiomIRISRibbon.Utility.SaveTempFile(_versionid + "X");
                        System.IO.File.Copy(strFileAttached, filenamecopy); // Create a copy

                        // Now save the copied file - change this to always save as the version name
                        Globals.ThisAddIn.ProcessingUpdate("Save To SalesForce");
                        string vfilename = _versionName.Replace(" ", "_") + ".docx";
                        dr = AxiomIRISRibbon.Utility.HandleData(_d.UpdateFile(_attachmentid, vfilename, filenamecopy));
                    }

                    Globals.ThisAddIn.AddSaveHandler(); // add it back in
                    Globals.ThisAddIn.ProcessingStop("End");

                }
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error saving contract");
                return false;
            }
        }

    }
}
