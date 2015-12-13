﻿using System;
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
        ~CompareSideBar()
        {
            //  System.Runtime.InteropServices.Marshal.ReleaseComObject(csb);
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

                //DataReturn dr = AxiomIRISRibbon.Utility.HandleData(_d.GetTemplates(true));
                DataReturn dr = AxiomIRISRibbon.Utility.HandleData(_d.GetAgreementTemplates(true));

                if (!dr.success) return;

                DataTable dt = dr.dt;
                cbTemplates.Items.Clear();

                RadComboBoxItem i;

                // RadComboBoxItem selected = null;
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
                //Logger.Log(ex, "Clone");
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
                    wordAttachment.TrackRevisions = false;
                    wordAttachment.ShowRevisions = false;
                    //wordAttachment.ActiveWindow.View.RevisionsFilter.Markup = Word.WdRevisionsMarkup.wdRevisionsMarkupNone;
                    wordAttachment.AcceptAllRevisions();

                    // To unlock Clauses
                    /*                   for (int i = 1; i <= wordAttachment.ContentControls.Count; i++)
                                       {
                                           wordAttachment.ContentControls[i].LockContents = false;
                                           wordAttachment.ContentControls[i].LockContentControl = false;

                                       }

                   */

                    /*   object objAttachment = _fileName;
                      wordAttachment = app.Documents.Open(ref objAttachment, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                        ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                       ref missing, ref missing);*/
                    /*  wordAttachment = app.Documents.Open(ref objAttachment, ref missing,true, ref missing, ref missing, ref missing, ref missing, ref missing,
                 ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing);
                    wordTemplate.ActiveWindow.View.ShowRevisionsAndComments = false;
                      //Compare
                 Globals.ThisAddIn.AddDocId(wordAttachment, "Compare", "");*/


                    object objTemplate = fileTemplate;
                    wordTemplate = app.Documents.Open(ref objTemplate, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);


                    // To unlock Clauses
                    /*                for (int i = 1; i <= wordTemplate.ContentControls.Count; i++)
                                    {
                                        wordTemplate.ContentControls[i].LockContents = false;
                                        wordTemplate.ContentControls[i].LockContentControl = false;

                                    }
                */
                    // Code to remove   the document modified by -- This code will reset all the properties in the document

                    wordTemplate.RemoveDocumentInformation(Microsoft.Office.Interop.Word.WdRemoveDocInfoType.wdRDIDocumentProperties);

                    // End Code

                    //Compare
                    // Globals.ThisAddIn.AddDocId(wordTemplate, "Compare", "");
                    Globals.ThisAddIn.AddDocId(wordTemplate, "Contract", "", "Compare");
                    wordTemplate.ActiveWindow.View.ShowRevisionsAndComments = false;
                    //wordTemplate.ActiveWindow.View.RevisionsFilter.Markup = Word.WdRevisionsMarkup.wdRevisionsMarkupNone;
                    wordTemplate.TrackRevisions = true;
                    wordTemplate.ShowRevisions = false;
                    wordTemplate.AcceptAllRevisions();


                    //added below lines to close the open file before opening the split screen
                    foreach (Word.Document d in app.Documents)
                    {
                        //d.ActiveWindow.View.RevisionsFilter.Markup = Word.WdRevisionsMarkup.wdRevisionsMarkupNone;
                        if (d.FullName != fileTemplate)
                        {

                            object saveOption = Word.WdSaveOptions.wdDoNotSaveChanges;
                            object originalFormat = Word.WdOriginalFormat.wdOriginalDocumentFormat;
                            object routeDocument = false;
                            ((Word._Document)d).Close(ref saveOption, ref originalFormat, ref routeDocument);
                        }
                    }


                    /*
                    object o = wordTemplate;
                    wordTemplate.Windows.CompareSideBySideWith(ref o);*/

                    //  Compare code
                    wordTemplate.Compare(_fileName, missing, Word.WdCompareTarget.wdCompareTargetNew, true, true, false, false, false);
                    app.ActiveWindow.View.SplitSpecial = Word.WdSpecialPane.wdPaneRevisionsVert;
                    app.ActiveWindow.ShowSourceDocuments = Word.WdShowSourceDocuments.wdShowSourceDocumentsOriginal;
                    //app.ActiveWindow.View.RevisionsFilter.Markup = 0;
                    app.ActiveWindow.View.RevisionsFilter.Markup = Word.WdRevisionsMarkup.wdRevisionsMarkupNone;
                    app.Activate();
                    // Russel Dec11 - add in the Doc Id to the comparison doc
                    Globals.ThisAddIn.AddDocId(app.ActiveDocument, "Contract", "", "Compare");

                    // close the temp files
                    var docTemplateClose = (Word._Document)wordTemplate;
                    docTemplateClose.Close(SaveChanges: false);


                    // System.Runtime.InteropServices.Marshal.ReleaseComObject(newdoc);
                    //  docclose = (Microsoft.Office.Interop.Word._Document)olddoc;
                    //  docclose.Close(SaveChanges: false);
                    //  System.Runtime.InteropServices.Marshal.ReleaseComObject(olddoc);

                    //  wordTemplate.Activate();
                    //End Compare
                    Globals.Ribbons.Ribbon1.CloseWindows();

                }
                else
                {
                    MessageBox.Show("Please select one template");

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error generating Compare");
                //Logger.Log(ex, "Clone");
            }
        }
        /*
          private void btnCompare_Click(object sender, RoutedEventArgs e)
        {ViewSideBySide();}
         Private void ViewSideBySide(){
            try
            {
                btnCompare.IsEnabled = false;
                object missing = System.Reflection.Missing.Value;

                if (this.cbTemplates.SelectedItem != null)
                {

                    string TemplateId = ((RadComboBoxItem)(this.cbTemplates.SelectedItem)).Tag.ToString();
                    string TemplateName = ((RadComboBoxItem)(this.cbTemplates.SelectedItem)).Content.ToString();
                    //Word.Document tempDoc;

                    DataReturn drAttachemnts = AxiomIRISRibbon.Utility.HandleData(_d.GetTemplateAttach(TemplateId));
                    if (!drAttachemnts.success) return;


                    DataTable dtAttachments = drAttachemnts.dt;
                    string fileTemplate = "";
                    foreach (DataRow rw in dtAttachments.Rows)
                    {
                        byte[] toBytes = Convert.FromBase64String(rw["body"].ToString());
                       // fileTemplate = _d.GetTempFilePath(rw["Id"].ToString() + "_" + rw["Name"].ToString());
                        fileTemplate = _d.GetTempFilePath("Template View");
                        File.WriteAllBytes(fileTemplate, toBytes);
                        //   _source = app.Documents.Open(filename);


                    }

                    Word.Document wordTemplate;
                    Word.Document wordAttachment;
                    Word.Application app = Globals.ThisAddIn.Application;

                    wordAttachment = Globals.ThisAddIn.Application.Documents.Add(_fileName);
                    wordAttachment.TrackRevisions = false;
                    wordAttachment.ShowRevisions = false;
                    wordAttachment.AcceptAllRevisions();

                    //// To unlock Clauses
                    //                  for (int i = 1; i <= wordAttachment.ContentControls.Count; i++)
                    //                   {
                    //                       wordAttachment.ContentControls[i].LockContents = false;
                    //                       wordAttachment.ContentControls[i].LockContentControl = false;

                    //                   }

               

                     object objAttachment = _fileName;
                      wordAttachment = app.Documents.Open(ref objAttachment, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                        ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                       ref missing, ref missing);
                      wordAttachment = app.Documents.Open(ref objAttachment, ref missing,true, ref missing, ref missing, ref missing, ref missing, ref missing,
                 ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing);
                    wordTemplate.ActiveWindow.View.ShowRevisionsAndComments = false;
                      //Compare
                 Globals.ThisAddIn.AddDocId(wordAttachment, "Compare", "");


                    object objTemplate = fileTemplate;
                    wordTemplate = app.Documents.Open(ref objTemplate, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);


                    //// To unlock Clauses
                    //for (int i = 1; i <= wordTemplate.ContentControls.Count; i++)
                    //{
                    //    wordTemplate.ContentControls[i].LockContents = false;
                    //    wordTemplate.ContentControls[i].LockContentControl = false;

                    //}


                    //Compare
                   // Globals.ThisAddIn.AddDocId(wordTemplate, "Compare", "");
                    Globals.ThisAddIn.AddDocId(wordTemplate, "Contract",  "","Compare");
                    wordTemplate.ActiveWindow.View.ShowRevisionsAndComments = false;
                    wordTemplate.TrackRevisions = true;
                    wordTemplate.ShowRevisions = false;
                    wordTemplate.AcceptAllRevisions();


                    //added below lines to close the open file before opening the split screen
                    Microsoft.Office.Interop.Word.Application WordObj;
                    WordObj = (Microsoft.Office.Interop.Word.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Word.Application");
                    foreach (Word.Document d in WordObj.Documents)
                    {
                        if (d.FullName == fileTemplate)
                        {
                            //nothing
                        }
                        else
                        {
                            //
                            object saveOption = Word.WdSaveOptions.wdDoNotSaveChanges;
                            object originalFormat = Word.WdOriginalFormat.wdOriginalDocumentFormat;
                            object routeDocument = false;
                            d.Close(ref saveOption, ref originalFormat, ref routeDocument);
                        }
                    }

                    
                    
                    //object o = wordTemplate;
                    //wordTemplate.Windows.CompareSideBySideWith(ref o);

                    //  Compare code
                    wordTemplate.Compare(_fileName, missing, Word.WdCompareTarget.wdCompareTargetNew, true, true, false, false, false);
                    app.ActiveWindow.View.SplitSpecial = Word.WdSpecialPane.wdPaneRevisionsVert;
                    app.ActiveWindow.ShowSourceDocuments = Word.WdShowSourceDocuments.wdShowSourceDocumentsOriginal;
                    app.ActiveWindow.View.RevisionsFilter.Markup = 0;
                    app.Activate();

                    // close the temp files
                    var docTemplateClose = (Word._Document)wordTemplate;
                    docTemplateClose.Close(SaveChanges: false);
                    var docAttachmentClose = (Word._Document)wordAttachment;
                    docAttachmentClose.Close(SaveChanges: false);


                    // System.Runtime.InteropServices.Marshal.ReleaseComObject(newdoc);
                    //  docclose = (Microsoft.Office.Interop.Word._Document)olddoc;
                    //  docclose.Close(SaveChanges: false);
                    //  System.Runtime.InteropServices.Marshal.ReleaseComObject(olddoc);

                    //  wordTemplate.Activate();
                    //End Compare
                    Globals.Ribbons.Ribbon1.CloseWindows();

                }
                else
                {
                    MessageBox.Show("Please select one template");

                }
            }
            catch (Exception ex)
            {
                //Logger.Log(ex, "Clone");
            }
        }
        */
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
