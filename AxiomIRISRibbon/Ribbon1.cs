using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using System.Data;
using System.Windows;
using Telerik.Windows.Controls;
using Word = Microsoft.Office.Interop.Word;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using AxiomIRISRibbon.SForceEdit;
using System.IO;
using System.Xml;
using System.Windows.Forms;
using AxiomIRISRibbon.ContractEdit;
using System.Diagnostics;

namespace AxiomIRISRibbon
{
    public partial class Axiom
    {

        int _sfcount = 0;
        private Data D;
       private string Id;
       private bool IsTemplate;
       private DataTable DTTemplate;
       private Word.Document Doc;
       private string _versionid;
       private Data _d;

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {            
            gpAdmin.Visible = false;
            gpData.Visible = false;
            gpDraft.Visible = false;
            btnLogout.Enabled = false;
            btnLogin.Enabled = true;
            btnTrack.Visible = false;

            btnReports.Enabled = false;

            gpDraft.Visible = false;
            gpIrisTrack.Visible = false;
        }

        public void Activate()
        {
            try
            {
                this.RibbonUI.ActivateTabMso("TabAddIns");
            }
            catch(Exception ex){

            }
        }


        public bool isUserAdmin(Data d){
            string profile = d.GetUserProfile();
            JToken s = Globals.ThisAddIn.GetSettings().GetGeneralSetting("AdminMenu");
            if(s!=null){
                if(s.Type == JTokenType.Array){
                    foreach(string t in (JArray)s){
                        if (t == profile) return true;
                    }
                }
                if (s.Type == JTokenType.String)
                {
                        if (s.ToString() == profile) return true;                    
                }
            }

            // Default so that these profiles always get them
            if (profile == "System Administrator" || profile == "Axiom Admin") return true;
            
            return false;
        }

        public void LoginOK()
        {
            btnLogout.Enabled = true;
            btnLogin.Enabled = true;
            btnReports.Enabled = true;

            Data d = Globals.ThisAddIn.getData();
            d.CheckTableNames();

            if (Globals.ThisAddIn.getData().HasLibraryObjects())
            {
                // for now just don't switch on the drafting menu - will want to add that back for the demo
                // gpDraft.Visible = true;

                JToken s = Globals.ThisAddIn.GetSettings().GetGeneralSetting("DraftMenu");
                if (s != null)
                {
                    if (s.ToString() == "On")
                    {
                        gpDraft.Visible = true;
                        RefreshTemplatesList();
                    }
                }

                if (isUserAdmin(d))
                {
                    gpAdmin.Visible = true;
                }
                else
                {
                    gpAdmin.Visible = false;
                }


            }

            //Check the instance has the objects that are in the toolbar
            //if not hide them

            btn1.Visible = false;
            btn2.Visible = false;
            btn3.Visible = false;
            btn4.Visible = false;
            btn5.Visible = false;

            // from the settings
            string toplevel = Globals.ThisAddIn.GetSettings("", "TopLevelSObjects");
            if (toplevel != "")
            {

                // CANT ADD BUTTONS TO THE GROUP AT RUNTIME! ARG
                // SO - set up 10 buttons and assign them to the objects in order and 
                // update the label and icon and make visible

                // RibbonGroup gp = Globals.Factory.GetRibbonFactory().CreateRibbonGroup();
                // gp.Label = "Data";
                // gp.Name = "gpData1";

                gpData.Visible = true;
                //Code PES
                gpIrisTrack.Visible = true;
                // End PES

                btn1.Visible = false;
                btn2.Visible = false;
                btn3.Visible = false;
                btn4.Visible = false;
                btn5.Visible = false;

                string[] tl = toplevel.Split('|');
                int btnnumber = 1;
                foreach (string tlentry in tl)
                {
                    if (btnnumber < 6)
                    {
                        string[] tle = tlentry.Split(':');
                        string name = tle[0];
                        string label = tle[1];
                        string icon = tle[2];

                        RibbonButton btn = null;

                        // yeh there will be a better way to do this - but can't see the items collection
                        if (btnnumber == 1)
                        {
                            btn = this.btn1;
                        }
                        else if (btnnumber == 2)
                        {
                            btn = this.btn2;
                        }
                        else if (btnnumber == 3)
                        {
                            btn = this.btn3;
                        }
                        else if (btnnumber == 4)
                        {
                            btn = this.btn4;
                        }
                        else if (btnnumber == 5)
                        {
                            btn = this.btn5;
                        }


                        btn.Label = label;
                        btn.Name = "btn" + btnnumber.ToString();
                        btn.Tag = name;
                        btn.ScreenTip = "Select to open Editor";
                        btn.SuperTip = "";

                        if (label != "")
                        {
                            try
                            {
                                System.Reflection.Assembly asm = System.Reflection.Assembly.GetExecutingAssembly();
                                string resourceName = asm.GetName().Name + ".Properties.Resources";
                                System.Resources.ResourceManager rm = new System.Resources.ResourceManager(resourceName, asm);
                                System.Drawing.Bitmap bmp = (System.Drawing.Bitmap)rm.GetObject(icon);
                                btn.Image = bmp;
                            }
                            catch (Exception)
                            {
                                // will just use the default
                            }
                        }

                        btn.ShowImage = true;
                        btn.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
                        btn.Visible = true;
                        // group1.Items.Add(btn);

                        btnnumber++;
                    }
                }
            }


            if (Globals.ThisAddIn.getDebug())
            {
                btnTrack.Visible = true;
            }

        }

        public void Logout(){
            Globals.ThisAddIn.getData().Logout();

            CloseWindows();
            Globals.ThisAddIn.HideWindows();

            // shut down all the edit windows


            gpAdmin.Visible = false;
            gpData.Visible = false;
            gpDraft.Visible = false;
            btnLogout.Enabled = false;
            btnLogin.Enabled = true;

            btnReports.Enabled = false;

            //Code PES
            gpIrisTrack.Visible = false;
            // End PES
        }


        public void RefreshTemplatesList(){
            Data d = Globals.ThisAddIn.getData();
           
            DataReturn dr = Utility.HandleData(d.GetTemplates(true));
            if (!dr.success) return;

            gContracts.Items.Clear();

            DataTable contracts = dr.dt;
            foreach (DataRow r in contracts.Rows)
            {
                RibbonDropDownItem i = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
                i.Label = r["Name"].ToString();
                i.ScreenTip = "Select to create an instance of this contract";
                i.SuperTip = r["Description__c"].ToString();
                i.Image = new System.Drawing.Bitmap(AxiomIRISRibbon.Properties.Resources.contract);
                i.Tag = r["Id"].ToString() + "|" + r["PlaybookLink__c"].ToString();
                gContracts.Items.Add(i);
            }
        }


        private void btnLogin_Click(object sender, RibbonControlEventArgs e)
        {
            CloseWindows();
            Globals.ThisAddIn.HideWindows();
            Globals.ThisAddIn.OpenLogin();
        }


        private void btnTemplate_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.OpenTemplate();
        }

        private void btnNewTemplate_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.OpenTemplate().NewTemplate();
        }

        private void btnNewFromExsisting_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.OpenTemplate().NewTemplate();
        }

        private void btnClauses_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.OpenClause(true,true);
        }

        private void btnElement_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.OpenElement();
        }

        private void gContracts_Click(object sender, RibbonControlEventArgs e)
        {                       
            Contract axC = new Contract();

            string tag = Convert.ToString(((RibbonGallery)sender).SelectedItem.Tag);
            string[] atag = tag.Split('|');
            string TemplateId = atag[0];
            string TemplatePlaybookLink = "";
            if (atag.Length > 1)
            {
                TemplatePlaybookLink = atag[1];
            }

            axC.Open("", TemplateId, Convert.ToString(((RibbonGallery)sender).SelectedItem.Label), TemplatePlaybookLink);                       
        }

        private void btnLogout_Click(object sender, RibbonControlEventArgs e)
        {
            this.Logout();
        }

        private void btnOpenContract_Click(object sender, RibbonControlEventArgs e)
        {
           Globals.ThisAddIn.OpenContract();
        }

        private void btnConcepts_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.OpenConcept();
        }

        private void btnBlankTemplate_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.Application.Documents.Add();
            Globals.ThisAddIn.OpenTemplate().NewTemplate();
        }

        private void btnNewClause_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.OpenContract().NewContract();
        }

        private void btnBlankClause_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.Application.Documents.Add();
            Globals.ThisAddIn.OpenContract().NewContract();
        }

        private void btnSendForApproval_Click(object sender, RibbonControlEventArgs e)
        {
            //Save a version of the currenct document and open outlook and attach
            Globals.ThisAddIn.GetTaskPaneControlContract().SaveAndSendApproval("", "", "");
        }

        public void Approval(bool approval)
        {
            if (approval)
            {
                btnSendForApproval.Enabled = true;
                btnSendForNeg.Enabled = false;
            }
            else
            {
                btnSendForApproval.Enabled = false;
                btnSendForNeg.Enabled = true;
            }

        }

        private void btnSendForNeg_Click(object sender, RibbonControlEventArgs e)
        {
            //Save a version of the currenct document and open outlook and attach
            Globals.ThisAddIn.GetTaskPaneControlContract().SaveAndSendNeg();
        }





        private void btnDataEdit_Click(object sender, RibbonControlEventArgs e)
        {
            RibbonButton b = (RibbonButton)sender;
            string tag = b.Tag.ToString();

            Globals.ThisAddIn.CloseEditWindows();
            Globals.ThisAddIn.OpenEditWindow(tag);
        }




        public void CloseWindows()
        {
            Globals.ThisAddIn.CloseEditWindows();
            Globals.ThisAddIn.CloseEditZoomWindows();
        }

        public void SFDebug(string desc,string sql)
        {
            if (btnTrack.Visible)
            {
                _sfcount++;
                lbSFCount.Label = _sfcount.ToString();

                RibbonDropDownItem i = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
                i.Label = _sfcount.ToString() + " " + desc;    
                i.Tag = sql;
                gSFDebug.Items.Add(i);
            }

        }
        public void SFDebug(string desc)
        {
            if (Globals.ThisAddIn.getDebug())
            {
                _sfcount++;
                lbSFCount.Label = _sfcount.ToString();
                RibbonDropDownItem i = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
                i.Label = _sfcount.ToString() + " " + desc;                
                gSFDebug.Items.Add(i);
                gSFDebug.SelectedItem = i;
            }

        }

        private void gSFDebug_Click(object sender, RibbonControlEventArgs e)
        {
            System.Windows.Forms.MessageBox.Show(Convert.ToString(((RibbonGallery)sender).SelectedItem.Label) + "\n" + Convert.ToString(((RibbonGallery)sender).SelectedItem.Tag));            
        }



        private void btnLoginSSO_Click(object sender, RibbonControlEventArgs e)
        {
            CloseWindows();
            Globals.ThisAddIn.HideWindows();
            Globals.ThisAddIn.OpenLoginSSO().Login(null);
        }

        private void btnSettings_Click(object sender, RibbonControlEventArgs e)
        {
            CloseWindows();
            Globals.ThisAddIn.HideWindows();
            Globals.ThisAddIn.OpenLocalSettings();
        }

        private void sbtnLoginSSO_Click(object sender, RibbonControlEventArgs e)
        {
            CloseWindows();
            Globals.ThisAddIn.HideWindows();
            Globals.ThisAddIn.OpenLoginSSO().Login(null);
        }

        private void btnLoginDev_Click(object sender, RibbonControlEventArgs e)
        {
            CloseWindows();
            Globals.ThisAddIn.HideWindows();
            Globals.ThisAddIn.OpenLoginSSO().Login(LocalSettings.Instances.Dev);
        }

        private void btnLoginIT_Click(object sender, RibbonControlEventArgs e)
        {
            CloseWindows();
            Globals.ThisAddIn.HideWindows();
            Globals.ThisAddIn.OpenLoginSSO().Login(LocalSettings.Instances.IT);
        }

        private void btnLoginUAT_Click(object sender, RibbonControlEventArgs e)
        {
            CloseWindows();
            Globals.ThisAddIn.HideWindows();
            Globals.ThisAddIn.OpenLoginSSO().Login(LocalSettings.Instances.UAT);
        }

        private void btnLoginProd_Click(object sender, RibbonControlEventArgs e)
        {
            CloseWindows();
            Globals.ThisAddIn.HideWindows();
            Globals.ThisAddIn.OpenLoginSSO().Login(LocalSettings.Instances.Prod);
        }

        private void btnReports_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.OpenReports();
        }

        private void btnAbout_Click(object sender, RibbonControlEventArgs e)
        {
            CloseWindows();
            Globals.ThisAddIn.HideWindows();
            Globals.ThisAddIn.OpenAbout();
        }
             
        //Code PES
        private void btnTrack_DialogLauncherClick(object sender, RibbonControlEventArgs e)
        {
       
           CompareAmendment.TrackDocument();
           // CompareAmendment.TrackDocumentOld();
        }
        private void btnExportToWord_Click(object sender, RibbonControlEventArgs e)
        {

            if (this.Id == null) this.Id = Globals.ThisAddIn.GetCurrentDocId();
            this.IsTemplate = true;
            //if (this.Id != "")
            //{
            //    this.IsTemplate = true;
            //    //GetTemplateDetails();

            //}
            // so whats the plan! 
            // for now don't make it human readable i.e. not a marked up doc - going to be too clumsy, might come back to
            // so leave in the controls and write the metadata and the non-shown clauses to custom xml parts

            // leave in the ids as they are - import will sort them out
            try
            {
                Globals.ThisAddIn.ProcessingStart("Import Document");

                if (this.IsTemplate)
                {

                    Globals.ThisAddIn.ProcessingUpdate("Copy Document");
                    Word.Document template = Globals.ThisAddIn.Application.ActiveDocument;
                    Word.Document export = Globals.ThisAddIn.Application.Documents.Add();

                    Word.Range source = template.Range(template.Content.Start, template.Content.End);
                    export.Range(export.Content.Start).InsertXML(source.WordOpenXML);

                    Globals.ThisAddIn.ProcessingUpdate("Update DocumentDocument Id");
                    //Globals.ThisAddIn.AddDocId(export, "ExportTemplate", this.Id);

                    // now get the meta data and store it as custom xml parts
                    Globals.ThisAddIn.ProcessingUpdate("Add Data to Dataset");
                    DataSet ds = new DataSet();
                    ds.Namespace = "http://www.axiomlaw.com/irisribbon";
                    //Globals.ThisAddIn.ProcessingUpdate("Template");
                    //ds.Tables.Add(this.DTTemplate);
                    //Globals.ThisAddIn.ProcessingUpdate("Clause");
                    //ds.Tables.Add(this.DTClause);
                    //Globals.ThisAddIn.ProcessingUpdate("ClauseXML");
                    //ds.Tables.Add(this.DTClauseXML);
                    //Globals.ThisAddIn.ProcessingUpdate("Elements");
                    //ds.Tables.Add(this.DTElement);

                    Globals.ThisAddIn.ProcessingUpdate("Serialise Data to XML");
                    string xmldata = "";
                    using (StringWriter stringWriter = new StringWriter())
                    {
                        ds.WriteXml(new XmlTextWriter(stringWriter));
                        xmldata = stringWriter.ToString();
                    };

                    xmldata = AxiomIRISRibbon.Utility.CleanUpXML(xmldata);

                    Globals.ThisAddIn.ProcessingUpdate("Save as XML Part");

                    Microsoft.Office.Core.CustomXMLPart data = export.CustomXMLParts.Add(xmldata);

                    export.Activate();

                    Globals.ThisAddIn.ProcessingUpdate("Save As ...");

                    SaveFileDialog dlg = new SaveFileDialog();
                    dlg.Title = "MyTitle";
                    dlg.Filter = "Word Document (*.doc;*.docx;*.docm)|*.doc;*.docx;*.docx";
                    dlg.FileName = "ExportDocument-" + dlg.Title.Replace(" ", "");
                    dlg.ShowDialog();
                    export.SaveAs2(dlg.FileName,ReadOnlyRecommended:true);
                    Globals.ThisAddIn.ProcessingStop("Finished");

                }
            }
            catch (Exception ex)
            {
                string message = "Sorry there has been an error - " + ex.Message;
                if (ex.InnerException != null) message += " " + ex.InnerException.Message;
                System.Windows.Forms.MessageBox.Show(message);
                // Globals.ThisAddIn.ProcessingStop("Finished");
            }

        }

        /*

        private void btnExportToPDF_Click(object sender, RibbonControlEventArgs e)
        {
            object oMissing = System.Reflection.Missing.Value;

            Word.Document template = Globals.ThisAddIn.Application.ActiveDocument;
            Word.Document export = Globals.ThisAddIn.Application.Documents.Add();

            Word.Range source = template.Range(template.Content.Start, template.Content.End);
            export.Range(export.Content.Start).InsertXML(source.WordOpenXML);

            export.Activate();

            object fileFormat = Word.WdSaveFormat.wdFormatPDF;
            SaveFileDialog dlg = new SaveFileDialog();
            dlg.Title = "MyTitle";
            dlg.Filter = "Word Document (*.pdf)|*.pdf";
            dlg.FileName = "ExportTemplate-" + dlg.Title.Replace(" ", "");
            dlg.ShowDialog();
            object outputFileName = dlg.FileName;
            export.SaveAs(ref outputFileName, ref fileFormat, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                ref oMissing, ref oMissing, ref oMissing);
            export.Activate();
            object saveChanges = Word.WdSaveOptions.wdDoNotSaveChanges;
            ((Microsoft.Office.Interop.Word._Document)export).Close(ref saveChanges, ref oMissing, ref oMissing);
            export = null;


        }

        */

        private void btnExportToPDF_Click(object sender, RibbonControlEventArgs e)
        {
            object oMissing = System.Reflection.Missing.Value;


            Word.Document template = Globals.ThisAddIn.Application.ActiveDocument;
            Word.Document export = Globals.ThisAddIn.Application.Documents.Add();

            Word.Range source = template.Range(template.Content.Start, template.Content.End);
            export.Range(export.Content.Start).InsertXML(source.WordOpenXML);

            Word.Shape logoWatermark = null;

            foreach (Word.Window item in export.Windows)
            {
                logoWatermark = item.Selection.Document.Content.Document.Shapes.AddTextEffect(
                            Microsoft.Office.Core.MsoPresetTextEffect.msoTextEffect1,
                            "Draft", "Arial", (float)60,
                            Microsoft.Office.Core.MsoTriState.msoTrue,
                            Microsoft.Office.Core.MsoTriState.msoFalse,
                            0, 0, ref oMissing);
                logoWatermark.Select(ref oMissing);
                logoWatermark.Fill.Visible = Microsoft.Office.Core.MsoTriState.msoTrue;
                logoWatermark.Line.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;
                logoWatermark.Fill.Solid();
                logoWatermark.Fill.Transparency = 0.2f;
                logoWatermark.Fill.ForeColor.RGB = (Int32)Word.WdColor.wdColorGray30;
                logoWatermark.RelativeHorizontalPosition = Word.WdRelativeHorizontalPosition.wdRelativeHorizontalPositionMargin;
                logoWatermark.RelativeVerticalPosition = Word.WdRelativeVerticalPosition.wdRelativeVerticalPositionMargin;
                logoWatermark.Left = (float)Word.WdShapePosition.wdShapeCenter;
                logoWatermark.Top = (float)Word.WdShapePosition.wdShapeCenter;
                logoWatermark.Height = Globals.ThisAddIn.Application.InchesToPoints(2.4f);
                logoWatermark.Width = Globals.ThisAddIn.Application.InchesToPoints(6f);
            }


            /// 
            /// 
            ///

            export.Activate();

            object fileFormat =Word.WdSaveFormat.wdFormatPDF;
            SaveFileDialog dlg = new SaveFileDialog();
            dlg.Title = "MyTitle";
            dlg.Filter = "Word Document (*.pdf)|*.pdf";
            dlg.FileName = "ExportTemplate-" + dlg.Title.Replace(" ", "");
            dlg.ShowDialog();
            object outputFileName = dlg.FileName;
            export.SaveAs(ref outputFileName, ref fileFormat, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                ref oMissing, ref oMissing, ref oMissing);
            //export.Activate();

            // Close the Word document, but leave the Word application open.
            // doc has to be cast to type _Document so that it will find the
            // correct Close method.                
            object saveChanges = Word.WdSaveOptions.wdDoNotSaveChanges;
            ((Word._Document)export).Close(ref saveChanges, ref oMissing, ref oMissing);
            export = null;
            //oWord.Quit();
            //System.Runtime.InteropServices.Marshal.ReleaseComObject(oWord); 
            //oWord = null;

        }
        private void btnAmend_Click(object sender, RibbonControlEventArgs e)
        {

            // CompareAmendment.FromRibbonToCreate();
            //  Id = Globals.ThisAddIn.Application.Documents.Add(attachmentid);
            // Globals.ThisAddIn.Application.Documents.Add(attachmentid) = Id;
            // System.Windows.Application.Current.Resources.GetValue("attachmentid") as Id;
            string attachmentId = Globals.ThisAddIn.GetAttachmentId();
            Id = attachmentId;
            Word.Document docTest = Globals.ThisAddIn.Application.ActiveDocument;
            if (this.Id == null) this.Id = Globals.ThisAddIn.GetCurrentDocId();
            if (this.Id != "")
            {
                _d = Globals.ThisAddIn.getData();
                DataReturn dr1 = AxiomIRISRibbon.Utility.HandleData(_d.AttachmentInfo(Id));
                if (!dr1.success) return;
                DataTable dtAllAttachments = dr1.dt;
                string _attachmentid = dtAllAttachments.Rows[0]["id"].ToString();
                string _filename = dtAllAttachments.Rows[0]["name"].ToString();
                string _parentId = dtAllAttachments.Rows[0]["ParentId"].ToString();

                CompareAmendment amend = new CompareAmendment();
                amend.Create(_attachmentid, _parentId, _filename);
                amend.Focus();
                amend.Show();
            }
            else
            {
                //_d = Globals.ThisAddIn.getData();
                //string id = Globals.ThisAddIn.GetCurrentDocId();
                //Word.Document doc = Globals.ThisAddIn.Application.Documents.Add(filename);
            }
        }

        //End PES

      
    }
}
