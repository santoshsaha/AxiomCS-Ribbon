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
using Telerik.Windows.Controls;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using System.Data;
using AxiomIRISRibbon.sfPartner;
using System.IO;
using System.ComponentModel;
using System.Windows.Threading;




namespace AxiomIRISRibbon.SForceEdit
{
    /// <summary>
    /// Interaction logic for Exsisting.xaml
    /// NEW File Added by PES
    /// </summary>
    public partial class Exsisting : RadWindow
    {

        static Microsoft.Office.Interop.Word.Application app;
        BackgroundWorker busyIndicatorBackgroundWorker;

        private Data _d;
        string _objname;
        string _id;
        string _name;
        public AxObject _parentObject;
        private static DataTable _dtAgreement;

        public Exsisting()
        {
            InitializeComponent();
            AxiomIRISRibbon.Utility.setTheme(this);

            _d = Globals.ThisAddIn.getData();

            app = Globals.ThisAddIn.Application;
        }

        private void ClauseRowDoubleClick(object sender, RoutedEventArgs e)
        {
            //Open();
        }


        public void Create(string objname, string id, string name, string templatename)
        {

            _objname = objname;
            _id = id;
            _name = name;
            DataTable emptyMatterDatatable = new DataTable();
            this.dgTemplates.ItemsSource = emptyMatterDatatable.DefaultView;
            dgTemplates.Focus();
        }
        public void btnReset_Click(object sender, RoutedEventArgs e)
        {
            CNID.Text = "";
            AgreemntNumber.Text = "";        

            this.dgTemplates.ItemsSource = _dtAgreement.DefaultView;
            dgTemplates.Focus();
        }

        protected void btnSearch_Click(object sender, RoutedEventArgs e)
        {

            string cnid;
            string agreementnumber;

            cnid = CNID.Text;
            agreementnumber = AgreemntNumber.Text;

            if (CNID.Text == "" && AgreemntNumber.Text == "")
            {
                MessageBoxResult result = MessageBox.Show("Please enter either Agreement Number or CNID");

            }
            else
            {

                try
                {
                    DataReturn dr = AxiomIRISRibbon.Utility.HandleData(_d.GetTemplateForsearch(agreementnumber, cnid));
                    
                    if (!dr.success) return;

                    DataTable dt = dr.dt;

                    if (dt.Rows.Count == 0)
                    {
                        MessageBox.Show("No eligible matter found");
                    }
                    this.dgTemplates.ItemsSource = dt.DefaultView;
                    dgTemplates.Focus();


                }
                catch (Exception ex)
                {
                    MessageBoxResult result = MessageBox.Show("Error text here", "Caption", MessageBoxButton.OK, MessageBoxImage.Error);
                }

            }
        }



        public void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
        //Code PES

        void busyIndicatorBackgroundWorker_DoWork(object sender, DoWorkEventArgs e, string strFromAgreementId)
        {
            BackgroundWorker worker = sender as BackgroundWorker;
            PerformCloning(worker, e, strFromAgreementId);
        }
        void busyIndicatorBackgroundWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {

            busyIndicatorBackgroundWorker.DoWork -= (obj, ev) => busyIndicatorBackgroundWorker_DoWork(obj, ev, null);
            busyIndicatorBackgroundWorker.RunWorkerCompleted -= busyIndicatorBackgroundWorker_RunWorkerCompleted;

            bsyIndc.IsBusy = false;
            bsyIndc.BusyContent = "";
            this.btnclone.IsEnabled = true;
            Globals.Ribbons.Ribbon1.CloseWindows();
            this.Close();
        }
     

        protected void btnClone_Click(object sender, RoutedEventArgs e)
        {

            this.btnclone.IsEnabled = false;
            bsyIndc.IsBusy = true;
            bsyIndc.BusyContent = "Cloning ...";
            busyIndicatorBackgroundWorker = new BackgroundWorker();
            string strFromAgreementId = string.Empty;
            if ((DataRowView)dgTemplates.SelectedItem == null)
            {
                
                this.btnclone.IsEnabled = true;
                bsyIndc.IsBusy = false;
                MessageBox.Show("Select an item.", "Alert");

            }
            else
            {
                DataRow dtr = ((DataRowView)dgTemplates.SelectedItem).Row;
                if (dtr != null)
                {
                    strFromAgreementId = dtr["Id"].ToString();
                    DataReturn dataReturnVersionRecord = _d.getLatestVersionDetails(strFromAgreementId);

                    if (dataReturnVersionRecord.dt.Rows.Count == 0)
                    {

                        this.btnclone.IsEnabled = true;
                        bsyIndc.IsBusy = false;

                        if (dataReturnVersionRecord.errormessage != "")
                        {
                            MessageBox.Show("Cloning cannot occur since selected Agreement does not have an available Version.", "Alert");
                        }
                        else
                        {
                            MessageBox.Show("Cloning cannot occur since selected Agreement does not have an available attachment.", "Alert");
                        }
                    }
                    else
                    {
                        busyIndicatorBackgroundWorker.DoWork += (obj, ev) => busyIndicatorBackgroundWorker_DoWork(obj, ev, strFromAgreementId);
                        busyIndicatorBackgroundWorker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(busyIndicatorBackgroundWorker_RunWorkerCompleted);
                        busyIndicatorBackgroundWorker.RunWorkerAsync();
                    }
                }


            }
        }
        protected void PerformCloning(BackgroundWorker worker, DoWorkEventArgs e, string strFromAgreementId)
        {
            try
            {

                if (string.IsNullOrEmpty(strFromAgreementId))
                {
                    MessageBox.Show("Select an item", "Alert");
                }
                else
                {
                    double dVersionNumber = 0;
                    string strToAgreementId, strFromVersionId = string.Empty, strTemplate = string.Empty;
                    strToAgreementId = _id;

                    DataRow allDr0, allDr1;
                    DataRow drSupersede, drSuperseded;

                    //Get version from 

                    DataReturn dr = AxiomIRISRibbon.Utility.HandleData(_d.GetAgreementsForVersion(strFromAgreementId, ""));

                    if (!dr.success) return;
                    if (dr.dt.Rows.Count == 0)
                    {
                        MessageBox.Show("Version not avilable in source Agreement", "Alert");
                    }
                    else
                    {
                        DataTable dtv = dr.dt;
                        allDr0 = dtv.NewRow();
                        allDr1 = dtv.NewRow();

                        foreach (DataRow rv in dtv.Rows)
                        {
                            strFromVersionId = rv["Id"].ToString();
                            strTemplate = Convert.ToString(rv["Template__c"]);
                        }
                        allDr0 = dtv.Rows[0];
                        allDr1.ItemArray = dtv.Rows[0].ItemArray.Clone() as object[];

                        //Get version to 
                        DataReturn drTo = AxiomIRISRibbon.Utility.HandleData(_d.GetAgreementsForVersion(strToAgreementId, ""));

                        double maxId;
                        if (drTo.dt.Rows.Count == 0)
                        {
                            maxId = 0;
                        }
                        else
                        {
                            dVersionNumber = Convert.ToDouble(drTo.dt.Rows[0]["version_number__c"]);
                            maxId = Convert.ToDouble(dVersionNumber + 1);
                        }

                        string VersionName = "Version " + (maxId).ToString();
                        string VersionNumber = maxId.ToString();

                        // Create Version 0 or lower version in To
                        DataReturn drCreatev0 = AxiomIRISRibbon.Utility.HandleData(_d.CreateVersion("", strToAgreementId, strTemplate, VersionName, VersionNumber, allDr0));
                        string newV0VersionId = drCreatev0.id;

                        _d.SaveVersionClause(strFromVersionId, newV0VersionId);

                        // Create Version 1 or lower version +1 in To
                        maxId = Convert.ToDouble(maxId + 1);
                        VersionName = "Version " + (maxId).ToString();
                        VersionNumber = maxId.ToString();

                        DataReturn drCreateV1 = AxiomIRISRibbon.Utility.HandleData(_d.CreateVersion("", strToAgreementId, strTemplate, VersionName, VersionNumber, allDr1));
                        string newV1VersionId = drCreateV1.id;  
                   
                        _d.SaveVersionClause(strFromVersionId, newV1VersionId);

                        //Code to update supersede and superseded by
                        //call query method
                        DataReturn dreturnSupersede = AxiomIRISRibbon.Utility.HandleData(_d.GetAgreementSupersedes(strFromAgreementId));
                        DataTable dtSupersede = new DataTable();
                        dtSupersede = dreturnSupersede.dt;
                        drSupersede = dtSupersede.NewRow();
                        foreach (DataRow r in dtSupersede.Rows)
                        {
                            drSupersede = r;
                        }
                        drSupersede["Superseded_By__c"] = strToAgreementId;

                        //call save method
                        _d.SaveMatter(drSupersede);

                        //call query method
                        DataReturn dreturnSuperseded = AxiomIRISRibbon.Utility.HandleData(_d.GetAgreementSupersedeby(strToAgreementId));
                        DataTable dtSuperseded = dreturnSuperseded.dt;
                        drSuperseded = dtSuperseded.NewRow();

                        //call save method
                        foreach (DataRow r in dtSuperseded.Rows)
                        {
                            drSuperseded = r;
                        }
                        drSuperseded["Supersedes__c"] = strFromAgreementId;

                        //call save method
                        _d.SaveMatter(drSuperseded);

                        //Create attachments in To
                        DataReturn drVersionAttachemnts = AxiomIRISRibbon.Utility.HandleData(_d.GetVersionAllAttachments(strFromVersionId));
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
                                _d.saveAttachmentstoSF(newV0VersionId, filename, body);
                                _d.saveAttachmentstoSF(newV1VersionId, filename, body);
                            }

                            //Get Attachments
                            DataReturn drAttachemnts = AxiomIRISRibbon.Utility.HandleData(_d.GetAllAttachments(newV1VersionId));
                            if (!drAttachemnts.success) return;
                            DataTable dtAllAttachments = drAttachemnts.dt;

                            //Open attachment with compare screeen
                            OpenAttachment(dtAllAttachments, newV1VersionId, strToAgreementId, strTemplate, VersionName, VersionNumber);

                        }
                    }
                }
            }
            catch (Exception ex)
            {
               
            }
           
        }

        private void OpenAttachment(DataTable dt, string versionid, string matterid, string templateid, string versionName, string versionNumber)
        {

            try
            {
                this.Dispatcher.BeginInvoke(DispatcherPriority.Background,
             new Action(
                 delegate
                 {
                     var res = from row in dt.AsEnumerable()
                               where
                               (row.Field<string>("Name").Contains(".doc") ||
                               row.Field<string>("ContentType").Contains("msword"))
                               select row;
                     if (res.Count() > 1)
                     {

                         AttachmentsView attTemp = new AttachmentsView();
                         attTemp.Create(dt, versionid, matterid, templateid, versionName, versionNumber);
                         attTemp.Show();
                         attTemp.Focus();
                     }
                     else
                     {

                         string attachmentid;
                         foreach (DataRow rw in dt.Rows)
                         {
                             if (rw["Name"].ToString().Contains(".doc"))
                             {
                                 byte[] toBytes = Convert.FromBase64String(rw["body"].ToString());
                                 string filename = _d.GetTempFilePath(rw["Id"].ToString() + "_" + rw["Name"].ToString());

                                 File.WriteAllBytes(filename, toBytes);

                                 
                                 Word.Document doc = Globals.ThisAddIn.Application.Documents.Add(filename);
                               
                                 attachmentid = rw["Id"].ToString();
                                 Globals.ThisAddIn.SaveAttachmentId(attachmentid);

                                 //Right Panel
                                 System.Windows.Forms.Integration.ElementHost elHost = new System.Windows.Forms.Integration.ElementHost();
                                 SForceEdit.CompareSideBar csb = new SForceEdit.CompareSideBar();
                                 csb.Create(filename, versionid, matterid, templateid, versionName, versionNumber, attachmentid);

                                 elHost.Child = csb;
                                 elHost.Dock = System.Windows.Forms.DockStyle.Fill;
                                 System.Windows.Forms.UserControl u = new System.Windows.Forms.UserControl();
                                 u.Controls.Add(elHost);
                                 Microsoft.Office.Tools.CustomTaskPane taskPaneValue = Globals.ThisAddIn.CustomTaskPanes.Add(u, "Axiom IRIS Compare", doc.ActiveWindow);
                                 taskPaneValue.Visible = true;
                                 taskPaneValue.Width = 400;


                             }
                         }
                     }
                 }));
            }
            catch (Exception ex)
            {
               
            }


        }



        //End Code PES
    }
}

