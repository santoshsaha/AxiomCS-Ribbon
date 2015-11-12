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
using AxiomIRISRibbon.sfPartner;
using System.IO;




namespace AxiomIRISRibbon.SForceEdit
{
    /// <summary>
    /// Interaction logic for NewFromTemplate.xaml
    /// </summary>
    public partial class Exsisting : RadWindow
    {

        static Microsoft.Office.Interop.Word.Application app;

        private Data _d;
        string _objname;
        string _id;
        string _name;
        private SForceEdit.SObjectDef _sDocumentObjectDef;
        public AxObject _parentObject;

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

            DataReturn dr = AxiomIRISRibbon.Utility.HandleData(_d.GetTemplatesFromExsisting(true));
            if (!dr.success) return;

            DataTable dt = dr.dt;
            dgTemplates.Items.Clear();

            this.dgTemplates.ItemsSource = dt.DefaultView;
            dgTemplates.Focus();
        }
        public void btnReset_Click(object sender, RoutedEventArgs e)
        {
            CNID.Text = "" ;
            AgreemntNumber.Text = "";
            DataReturn dr = AxiomIRISRibbon.Utility.HandleData(_d.GetTemplatesFromExsisting(true));
            if (!dr.success) return;

            DataTable dt = dr.dt;
            // dgTemplates.Items.Clear();
            this.dgTemplates.ItemsSource = dt.DefaultView;
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
               MessageBoxResult result = MessageBox.Show("Please enter either Agreemnt Number or CNID");

            }
            else
            {

                try
                {
                    DataReturn dr = AxiomIRISRibbon.Utility.HandleData(_d.GetTemplateForsearch(agreementnumber, cnid));
                    if (!dr.success) return;

                    DataTable dt = dr.dt;
                    //dgTemplates.Items.Clear();
                    this.dgTemplates.ItemsSource = dt.DefaultView;
                    dgTemplates.Focus();
                }
                catch (Exception ex)
                {
                    Logger.Log(ex, "btnSearch_Click");
                    MessageBoxResult result = MessageBox.Show("Error text here", "Caption", MessageBoxButton.OK, MessageBoxImage.Error);
                }

            }
        }



        public void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
        //Code PES
        protected void btnClone_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if ((DataRowView)dgTemplates.SelectedItem == null)
                {
                    MessageBox.Show("Select an item", "Alert");
                }
                else
                {
               //     bsyInd.IsBusy = true;
                //    bsyInd.BusyContent = "Cloning ...";
                    double dVersionNumber = 0;
                    string strFromAgreementId,strToAgreementId, strVersionId = string.Empty, strTemplate = string.Empty;
                    strToAgreementId=_id;

                    DataRow dtr = ((DataRowView)dgTemplates.SelectedItem).Row;
                  //  DataRow allDr;// = new DataRow();

                    strFromAgreementId = dtr["Id"].ToString();

                    //Get version from 
                    
                    DataReturn dr = AxiomIRISRibbon.Utility.HandleData(_d.GetAgreementLatestVersion(strFromAgreementId));
                    
                    if (!dr.success) return;

                    DataTable dt = dr.dt;
                   // allDr = dt.NewRow();

                    foreach (DataRow r in dt.Rows)
                    {
                        strVersionId = r["Id"].ToString();
                     //   dVersionNumber = Convert.ToDouble(r["version_number__c"]);
                        strTemplate = Convert.ToString(r["Template__c"]);
                    }

                    //Get version to 
                    DataReturn drFrom = AxiomIRISRibbon.Utility.HandleData(_d.GetAgreementLatestVersion(strToAgreementId));
                    DataTable dtrFrom = drFrom.dt;
                    foreach (DataRow rw in dtrFrom.Rows)
                    {
                        //strVersionId = rw["Id"].ToString();
                        dVersionNumber = Convert.ToDouble(rw["version_number__c"]);
                       // strTemplate = Convert.ToString(r["Template__c"]);
                    }
                    double maxId = Convert.ToDouble(dVersionNumber + 1);
                    string VersionName = "Version " + (maxId).ToString();
                    string VersionNumber = maxId.ToString();

                    // Create Version in TO
                    DataReturn drCreate = AxiomIRISRibbon.Utility.HandleData(_d.CreateVersion("", strToAgreementId, strTemplate, VersionName, VersionNumber));
                    string laterstVersionId = drCreate.id;

                    //Create attachments in To
                    DataReturn drVersionAttachemnts = AxiomIRISRibbon.Utility.HandleData(_d.GetVersionAllAttachments(strVersionId));
                    if (!drVersionAttachemnts.success) return;
                    DataTable dtAttachments = drVersionAttachemnts.dt;
                    string filename = "";
                    foreach (DataRow rw in dtAttachments.Rows)
                    {
                        filename = rw["Name"].ToString();
                        string body = rw["body"].ToString();
                        _d.saveAttachmentstoSF(laterstVersionId, filename, body);
                    }

                    //Get Attachments
                    DataReturn drAttachemnts = AxiomIRISRibbon.Utility.HandleData(_d.GetAllAttachments(laterstVersionId));
                    if (!drAttachemnts.success) return;
                    DataTable dtAllAttachments = drAttachemnts.dt;

                    //Open attachment with compare screeen
                    OpenAttachment(dtAllAttachments);
                    _sDocumentObjectDef = new SForceEdit.SObjectDef("Version__c");
                    Globals.Ribbons.Ribbon1.CloseWindows();

                }
            }
            catch (Exception ex)
            {
                Logger.Log(ex, "Clone");
            }
        }

        private  void OpenAttachment(DataTable dt)
        {
            foreach (DataRow rw in dt.Rows)
            {
                if (rw["Name"].ToString().Contains("Version"))
                {
                    byte[] toBytes = Convert.FromBase64String(rw["body"].ToString());
                    string filename = _d.GetTempFilePath(rw["Id"].ToString() + "_" + rw["Name"].ToString());             

                    File.WriteAllBytes(filename, toBytes);
                    // _source = app.Documents.Open(filename);

                 
                    Word.Document doc = Globals.ThisAddIn.Application.Documents.Add(filename);
                    //     Globals.ThisAddIn.Application.ActiveDocument.ActiveWindow.View.Type = Word.WdViewType.wdPrintView;
                    //     doc.Activate();

                    //Right Panel
                    System.Windows.Forms.Integration.ElementHost elHost = new System.Windows.Forms.Integration.ElementHost();
                    SForceEdit.CompareSideBar csb = new SForceEdit.CompareSideBar();
                    csb.Create(filename);

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

       

        //End Code PES
    }
}

