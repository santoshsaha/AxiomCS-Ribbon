﻿namespace AxiomIRISRibbon
{
    partial class Axiom : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Axiom()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.tbMain = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.btnLogin = this.Factory.CreateRibbonButton();
            this.sbtnLoginSSO = this.Factory.CreateRibbonSplitButton();
            this.btnLoginDev = this.Factory.CreateRibbonButton();
            this.btnLoginIT = this.Factory.CreateRibbonButton();
            this.btnLoginUAT = this.Factory.CreateRibbonButton();
            this.btnLoginProd = this.Factory.CreateRibbonButton();
            this.btnLoginSSO = this.Factory.CreateRibbonButton();
            this.btnLogout = this.Factory.CreateRibbonButton();
            this.gpData = this.Factory.CreateRibbonGroup();
            this.btn1 = this.Factory.CreateRibbonButton();
            this.btn2 = this.Factory.CreateRibbonButton();
            this.btn3 = this.Factory.CreateRibbonButton();
            this.btn4 = this.Factory.CreateRibbonButton();
            this.btn5 = this.Factory.CreateRibbonButton();
            this.gpAdmin = this.Factory.CreateRibbonGroup();
            this.btnTemplate = this.Factory.CreateRibbonSplitButton();
            this.btnNewTemplate = this.Factory.CreateRibbonButton();
            this.btnBlankTemplate = this.Factory.CreateRibbonButton();
            this.btnConcepts = this.Factory.CreateRibbonButton();
            this.btnClauses = this.Factory.CreateRibbonSplitButton();
            this.btnNewClause = this.Factory.CreateRibbonButton();
            this.btnBlankClause = this.Factory.CreateRibbonButton();
            this.btnElement = this.Factory.CreateRibbonSplitButton();
            this.gpDraft = this.Factory.CreateRibbonGroup();
            this.gContracts = this.Factory.CreateRibbonGallery();
            this.btnOpenContract = this.Factory.CreateRibbonButton();
            this.btnSendForApproval = this.Factory.CreateRibbonButton();
            this.btnSendForNeg = this.Factory.CreateRibbonButton();
            this.btnTrack = this.Factory.CreateRibbonGroup();
            this.lbSFCount = this.Factory.CreateRibbonLabel();
            this.lbSFLast = this.Factory.CreateRibbonLabel();
            this.gSFDebug = this.Factory.CreateRibbonGallery();
            this.gpIrisTrack = this.Factory.CreateRibbonGroup();
            this.btnRevertClause = this.Factory.CreateRibbonButton();
            this.btnExportToWord = this.Factory.CreateRibbonButton();
            this.btnExportToPDF = this.Factory.CreateRibbonButton();
            this.grpAmendment = this.Factory.CreateRibbonGroup();
            this.btnSync = this.Factory.CreateRibbonButton();
            this.btnAmend = this.Factory.CreateRibbonButton();
            this.group3 = this.Factory.CreateRibbonGroup();
            this.btnReports = this.Factory.CreateRibbonButton();
            this.btnSettings = this.Factory.CreateRibbonButton();
            this.btnAbout = this.Factory.CreateRibbonButton();
            this.btnNewFromExsisting = this.Factory.CreateRibbonButton();
            this.tbMain.SuspendLayout();
            this.group1.SuspendLayout();
            this.gpData.SuspendLayout();
            this.gpAdmin.SuspendLayout();
            this.gpDraft.SuspendLayout();
            this.btnTrack.SuspendLayout();
            this.gpIrisTrack.SuspendLayout();
            this.grpAmendment.SuspendLayout();
            this.group3.SuspendLayout();
            // 
            // tbMain
            // 
            this.tbMain.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tbMain.Groups.Add(this.group1);
            this.tbMain.Groups.Add(this.gpData);
            this.tbMain.Groups.Add(this.gpAdmin);
            this.tbMain.Groups.Add(this.gpDraft);
            this.tbMain.Groups.Add(this.btnTrack);
            this.tbMain.Groups.Add(this.gpIrisTrack);
            this.tbMain.Groups.Add(this.grpAmendment);
            this.tbMain.Groups.Add(this.group3);
            this.tbMain.Label = "AxiomIRIS";
            this.tbMain.Name = "tbMain";
            // 
            // group1
            // 
            this.group1.Items.Add(this.btnLogin);
            this.group1.Items.Add(this.sbtnLoginSSO);
            this.group1.Items.Add(this.btnLoginSSO);
            this.group1.Items.Add(this.btnLogout);
            this.group1.Label = "Connect";
            this.group1.Name = "group1";
            // 
            // btnLogin
            // 
            this.btnLogin.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnLogin.Image = global::AxiomIRISRibbon.Properties.Resources.Iris_Logo_Solo_Orange_40;
            this.btnLogin.Label = "Login";
            this.btnLogin.Name = "btnLogin";
            this.btnLogin.ShowImage = true;
            this.btnLogin.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnLogin_Click);
            // 
            // sbtnLoginSSO
            // 
            this.sbtnLoginSSO.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.sbtnLoginSSO.Image = global::AxiomIRISRibbon.Properties.Resources.Iris_Logo_Solo_Orange_40;
            this.sbtnLoginSSO.Items.Add(this.btnLoginDev);
            this.sbtnLoginSSO.Items.Add(this.btnLoginIT);
            this.sbtnLoginSSO.Items.Add(this.btnLoginUAT);
            this.sbtnLoginSSO.Items.Add(this.btnLoginProd);
            this.sbtnLoginSSO.ItemSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.sbtnLoginSSO.Label = "Login";
            this.sbtnLoginSSO.Name = "sbtnLoginSSO";
            this.sbtnLoginSSO.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.sbtnLoginSSO_Click);
            // 
            // btnLoginDev
            // 
            this.btnLoginDev.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnLoginDev.Image = global::AxiomIRISRibbon.Properties.Resources.asterix;
            this.btnLoginDev.Label = "Dev";
            this.btnLoginDev.Name = "btnLoginDev";
            this.btnLoginDev.ShowImage = true;
            this.btnLoginDev.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnLoginDev_Click);
            // 
            // btnLoginIT
            // 
            this.btnLoginIT.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnLoginIT.Image = global::AxiomIRISRibbon.Properties.Resources.asterix;
            this.btnLoginIT.Label = "IT";
            this.btnLoginIT.Name = "btnLoginIT";
            this.btnLoginIT.ShowImage = true;
            this.btnLoginIT.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnLoginIT_Click);
            // 
            // btnLoginUAT
            // 
            this.btnLoginUAT.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnLoginUAT.Image = global::AxiomIRISRibbon.Properties.Resources.asterix;
            this.btnLoginUAT.Label = "UAT";
            this.btnLoginUAT.Name = "btnLoginUAT";
            this.btnLoginUAT.ShowImage = true;
            this.btnLoginUAT.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnLoginUAT_Click);
            // 
            // btnLoginProd
            // 
            this.btnLoginProd.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnLoginProd.Image = global::AxiomIRISRibbon.Properties.Resources.asterix;
            this.btnLoginProd.Label = "Prod";
            this.btnLoginProd.Name = "btnLoginProd";
            this.btnLoginProd.ShowImage = true;
            this.btnLoginProd.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnLoginProd_Click);
            // 
            // btnLoginSSO
            // 
            this.btnLoginSSO.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnLoginSSO.Image = global::AxiomIRISRibbon.Properties.Resources.Iris_Logo_Solo_Orange_40;
            this.btnLoginSSO.Label = "Login";
            this.btnLoginSSO.Name = "btnLoginSSO";
            this.btnLoginSSO.ShowImage = true;
            this.btnLoginSSO.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnLoginSSO_Click);
            // 
            // btnLogout
            // 
            this.btnLogout.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnLogout.Image = global::AxiomIRISRibbon.Properties.Resources.signout;
            this.btnLogout.Label = "Logout";
            this.btnLogout.Name = "btnLogout";
            this.btnLogout.ShowImage = true;
            this.btnLogout.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnLogout_Click);
            // 
            // gpData
            // 
            this.gpData.Items.Add(this.btn1);
            this.gpData.Items.Add(this.btn2);
            this.gpData.Items.Add(this.btn3);
            this.gpData.Items.Add(this.btn4);
            this.gpData.Items.Add(this.btn5);
            this.gpData.Label = "Data";
            this.gpData.Name = "gpData";
            // 
            // btn1
            // 
            this.btn1.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btn1.Image = global::AxiomIRISRibbon.Properties.Resources.asterix;
            this.btn1.Label = "One";
            this.btn1.Name = "btn1";
            this.btn1.ShowImage = true;
            this.btn1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnDataEdit_Click);
            // 
            // btn2
            // 
            this.btn2.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btn2.Image = global::AxiomIRISRibbon.Properties.Resources.asterix;
            this.btn2.Label = "Two";
            this.btn2.Name = "btn2";
            this.btn2.ShowImage = true;
            this.btn2.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnDataEdit_Click);
            // 
            // btn3
            // 
            this.btn3.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btn3.Image = global::AxiomIRISRibbon.Properties.Resources.asterix;
            this.btn3.Label = "Three";
            this.btn3.Name = "btn3";
            this.btn3.ShowImage = true;
            this.btn3.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnDataEdit_Click);
            // 
            // btn4
            // 
            this.btn4.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btn4.Image = global::AxiomIRISRibbon.Properties.Resources.asterix;
            this.btn4.Label = "Four";
            this.btn4.Name = "btn4";
            this.btn4.ShowImage = true;
            this.btn4.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnDataEdit_Click);
            // 
            // btn5
            // 
            this.btn5.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btn5.Image = global::AxiomIRISRibbon.Properties.Resources.asterix;
            this.btn5.Label = "Five";
            this.btn5.Name = "btn5";
            this.btn5.ShowImage = true;
            this.btn5.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnDataEdit_Click);
            // 
            // gpAdmin
            // 
            this.gpAdmin.Items.Add(this.btnTemplate);
            this.gpAdmin.Items.Add(this.btnConcepts);
            this.gpAdmin.Items.Add(this.btnClauses);
            this.gpAdmin.Items.Add(this.btnElement);
            this.gpAdmin.Label = "Admin";
            this.gpAdmin.Name = "gpAdmin";
            // 
            // btnTemplate
            // 
            this.btnTemplate.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnTemplate.Image = global::AxiomIRISRibbon.Properties.Resources.document;
            this.btnTemplate.Items.Add(this.btnNewTemplate);
            this.btnTemplate.Items.Add(this.btnBlankTemplate);
            this.btnTemplate.ItemSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnTemplate.Label = "Templates";
            this.btnTemplate.Name = "btnTemplate";
            this.btnTemplate.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnTemplate_Click);
            // 
            // btnNewTemplate
            // 
            this.btnNewTemplate.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnNewTemplate.Image = global::AxiomIRISRibbon.Properties.Resources.document;
            this.btnNewTemplate.Label = "New From Current Document";
            this.btnNewTemplate.Name = "btnNewTemplate";
            this.btnNewTemplate.ShowImage = true;
            this.btnNewTemplate.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnNewTemplate_Click);
            // 
            // btnBlankTemplate
            // 
            this.btnBlankTemplate.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnBlankTemplate.Image = global::AxiomIRISRibbon.Properties.Resources.document;
            this.btnBlankTemplate.Label = "New Blank Template";
            this.btnBlankTemplate.Name = "btnBlankTemplate";
            this.btnBlankTemplate.ShowImage = true;
            this.btnBlankTemplate.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnBlankTemplate_Click);
            // 
            // btnConcepts
            // 
            this.btnConcepts.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnConcepts.Image = global::AxiomIRISRibbon.Properties.Resources.square;
            this.btnConcepts.Label = "Concepts";
            this.btnConcepts.Name = "btnConcepts";
            this.btnConcepts.ShowImage = true;
            this.btnConcepts.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnConcepts_Click);
            // 
            // btnClauses
            // 
            this.btnClauses.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnClauses.Image = global::AxiomIRISRibbon.Properties.Resources.clause;
            this.btnClauses.Items.Add(this.btnNewClause);
            this.btnClauses.Items.Add(this.btnBlankClause);
            this.btnClauses.Label = "Clauses";
            this.btnClauses.Name = "btnClauses";
            this.btnClauses.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnClauses_Click);
            // 
            // btnNewClause
            // 
            this.btnNewClause.Image = global::AxiomIRISRibbon.Properties.Resources.clause;
            this.btnNewClause.Label = "New From Current Document";
            this.btnNewClause.Name = "btnNewClause";
            this.btnNewClause.ShowImage = true;
            this.btnNewClause.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnNewClause_Click);
            // 
            // btnBlankClause
            // 
            this.btnBlankClause.Image = global::AxiomIRISRibbon.Properties.Resources.clause;
            this.btnBlankClause.Label = "New Blank Clause";
            this.btnBlankClause.Name = "btnBlankClause";
            this.btnBlankClause.ShowImage = true;
            this.btnBlankClause.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnBlankClause_Click);
            // 
            // btnElement
            // 
            this.btnElement.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnElement.Image = global::AxiomIRISRibbon.Properties.Resources.element;
            this.btnElement.Label = "Elements";
            this.btnElement.Name = "btnElement";
            this.btnElement.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnElement_Click);
            // 
            // gpDraft
            // 
            this.gpDraft.Items.Add(this.gContracts);
            this.gpDraft.Items.Add(this.btnOpenContract);
            this.gpDraft.Items.Add(this.btnSendForApproval);
            this.gpDraft.Items.Add(this.btnSendForNeg);
            this.gpDraft.Label = "Draft";
            this.gpDraft.Name = "gpDraft";
            // 
            // gContracts
            // 
            this.gContracts.ColumnCount = 1;
            this.gContracts.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.gContracts.Image = global::AxiomIRISRibbon.Properties.Resources.contract;
            this.gContracts.ItemImageSize = new System.Drawing.Size(32, 32);
            this.gContracts.Label = "New Contract";
            this.gContracts.Name = "gContracts";
            this.gContracts.RowCount = 3;
            this.gContracts.ShowImage = true;
            this.gContracts.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.gContracts_Click);
            // 
            // btnOpenContract
            // 
            this.btnOpenContract.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnOpenContract.Image = global::AxiomIRISRibbon.Properties.Resources.contract;
            this.btnOpenContract.Label = "Open Contract";
            this.btnOpenContract.Name = "btnOpenContract";
            this.btnOpenContract.ShowImage = true;
            this.btnOpenContract.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnOpenContract_Click);
            // 
            // btnSendForApproval
            // 
            this.btnSendForApproval.Enabled = false;
            this.btnSendForApproval.Image = global::AxiomIRISRibbon.Properties.Resources.sendmall;
            this.btnSendForApproval.Label = "Send For Approval";
            this.btnSendForApproval.Name = "btnSendForApproval";
            this.btnSendForApproval.ShowImage = true;
            this.btnSendForApproval.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSendForApproval_Click);
            // 
            // btnSendForNeg
            // 
            this.btnSendForNeg.Enabled = false;
            this.btnSendForNeg.Image = global::AxiomIRISRibbon.Properties.Resources.sendmall;
            this.btnSendForNeg.Label = "Send For Negotiating";
            this.btnSendForNeg.Name = "btnSendForNeg";
            this.btnSendForNeg.ShowImage = true;
            this.btnSendForNeg.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSendForNeg_Click);
            // 
            // btnTrack
            // 
            this.btnTrack.Items.Add(this.lbSFCount);
            this.btnTrack.Items.Add(this.lbSFLast);
            this.btnTrack.Items.Add(this.gSFDebug);
            this.btnTrack.Label = "Debug";
            this.btnTrack.Name = "btnTrack";
            this.btnTrack.DialogLauncherClick += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnTrack_DialogLauncherClick);
            // 
            // lbSFCount
            // 
            this.lbSFCount.Label = "0";
            this.lbSFCount.Name = "lbSFCount";
            // 
            // lbSFLast
            // 
            this.lbSFLast.Label = " ";
            this.lbSFLast.Name = "lbSFLast";
            // 
            // gSFDebug
            // 
            this.gSFDebug.ColumnCount = 1;
            this.gSFDebug.Label = "SF Calls";
            this.gSFDebug.Name = "gSFDebug";
            this.gSFDebug.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.gSFDebug_Click);
            // 
            // gpIrisTrack
            // 
            this.gpIrisTrack.Items.Add(this.btnRevertClause);
            this.gpIrisTrack.Items.Add(this.btnExportToWord);
            this.gpIrisTrack.Items.Add(this.btnExportToPDF);
            this.gpIrisTrack.Label = "Export";
            this.gpIrisTrack.Name = "gpIrisTrack";
            this.gpIrisTrack.Visible = false;
            // 
            // btnRevertClause
            // 
            this.btnRevertClause.Image = global::AxiomIRISRibbon.Properties.Resources.Iris_Logo_Solo_Orange_40;
            this.btnRevertClause.Label = "Revert Clause";
            this.btnRevertClause.Name = "btnRevertClause";
            this.btnRevertClause.ShowImage = true;
            this.btnRevertClause.Visible = false;
            this.btnRevertClause.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnRevertClause_Click);
            // 
            // btnExportToWord
            // 
            this.btnExportToWord.Image = global::AxiomIRISRibbon.Properties.Resources.Iris_Logo_Solo_Orange_40;
            this.btnExportToWord.Label = "Export To Word";
            this.btnExportToWord.Name = "btnExportToWord";
            this.btnExportToWord.ShowImage = true;
            this.btnExportToWord.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnExportToWord_Click);
            // 
            // btnExportToPDF
            // 
            this.btnExportToPDF.Image = global::AxiomIRISRibbon.Properties.Resources.Iris_Logo_Solo_Orange_40;
            this.btnExportToPDF.Label = "Export To PDF";
            this.btnExportToPDF.Name = "btnExportToPDF";
            this.btnExportToPDF.ShowImage = true;
            this.btnExportToPDF.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnExportToPDF_Click);
            // 
            // grpAmendment
            // 
            this.grpAmendment.Items.Add(this.btnSync);
            this.grpAmendment.Items.Add(this.btnAmend);
            this.grpAmendment.Label = "Amendment";
            this.grpAmendment.Name = "grpAmendment";
            this.grpAmendment.Visible = false;
            // 
            // btnSync
            // 
            this.btnSync.Image = global::AxiomIRISRibbon.Properties.Resources.Iris_Logo_Solo_Orange_40;
            this.btnSync.Label = "Sync";
            this.btnSync.Name = "btnSync";
            this.btnSync.ShowImage = true;
            this.btnSync.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnTrack_DialogLauncherClick);
            // 
            // btnAmend
            // 
            this.btnAmend.Image = global::AxiomIRISRibbon.Properties.Resources.Iris_Logo_Solo_Orange_40;
            this.btnAmend.Label = "Create";
            this.btnAmend.Name = "btnAmend";
            this.btnAmend.ShowImage = true;
            this.btnAmend.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAmend_Click);
            // 
            // group3
            // 
            this.group3.Items.Add(this.btnReports);
            this.group3.Items.Add(this.btnSettings);
            this.group3.Items.Add(this.btnAbout);
            this.group3.Label = "IRIS";
            this.group3.Name = "group3";
            // 
            // btnReports
            // 
            this.btnReports.Image = global::AxiomIRISRibbon.Properties.Resources.reports;
            this.btnReports.Label = "Reports";
            this.btnReports.Name = "btnReports";
            this.btnReports.ShowImage = true;
            this.btnReports.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnReports_Click);
            // 
            // btnSettings
            // 
            this.btnSettings.Image = global::AxiomIRISRibbon.Properties.Resources.cog;
            this.btnSettings.Label = "Settings";
            this.btnSettings.Name = "btnSettings";
            this.btnSettings.ShowImage = true;
            this.btnSettings.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSettings_Click);
            // 
            // btnAbout
            // 
            this.btnAbout.Image = global::AxiomIRISRibbon.Properties.Resources.Iris_Logo_Solo_Orange_40;
            this.btnAbout.Label = "About";
            this.btnAbout.Name = "btnAbout";
            this.btnAbout.ShowImage = true;
            this.btnAbout.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAbout_Click);
            // 
            // btnNewFromExsisting
            // 
            this.btnNewFromExsisting.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnNewFromExsisting.Image = global::AxiomIRISRibbon.Properties.Resources.document;
            this.btnNewFromExsisting.Label = "New From Current Document";
            this.btnNewFromExsisting.Name = "btnNewFromExsisting";
            this.btnNewFromExsisting.ShowImage = true;
            this.btnNewFromExsisting.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnNewFromExsisting_Click);
            // 
            // Axiom
            // 
            this.Name = "Axiom";
            this.RibbonType = "Microsoft.Word.Document";
            this.Tabs.Add(this.tbMain);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tbMain.ResumeLayout(false);
            this.tbMain.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.gpData.ResumeLayout(false);
            this.gpData.PerformLayout();
            this.gpAdmin.ResumeLayout(false);
            this.gpAdmin.PerformLayout();
            this.gpDraft.ResumeLayout(false);
            this.gpDraft.PerformLayout();
            this.btnTrack.ResumeLayout(false);
            this.btnTrack.PerformLayout();
            this.gpIrisTrack.ResumeLayout(false);
            this.gpIrisTrack.PerformLayout();
            this.grpAmendment.ResumeLayout(false);
            this.grpAmendment.PerformLayout();
            this.group3.ResumeLayout(false);
            this.group3.PerformLayout();

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tbMain;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnLogin;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup gpAdmin;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnLogout;
        internal Microsoft.Office.Tools.Ribbon.RibbonSplitButton btnTemplate;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnNewTemplate;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnNewFromExsisting;  //Jyoti
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnBlankTemplate;
        internal Microsoft.Office.Tools.Ribbon.RibbonSplitButton btnClauses;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnNewClause;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnBlankClause;
        internal Microsoft.Office.Tools.Ribbon.RibbonSplitButton btnElement;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup gpDraft;
        internal Microsoft.Office.Tools.Ribbon.RibbonGallery gContracts;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnOpenContract;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnConcepts;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSendForApproval;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSendForNeg;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup gpData;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn4;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup btnTrack;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel lbSFCount;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel lbSFLast;
        internal Microsoft.Office.Tools.Ribbon.RibbonGallery gSFDebug;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnLoginSSO;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSettings;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAbout;
        internal Microsoft.Office.Tools.Ribbon.RibbonSplitButton sbtnLoginSSO;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnLoginDev;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnLoginIT;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnLoginUAT;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnLoginProd;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnReports;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn5;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup gpIrisTrack;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSync;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAmend;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnExportToWord;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnExportToPDF;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnRevertClause;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpAmendment;
    }

    partial class ThisRibbonCollection
    {
        internal Axiom Ribbon1
        {
            get { return this.GetRibbon<Axiom>(); }
        }
    }
}
