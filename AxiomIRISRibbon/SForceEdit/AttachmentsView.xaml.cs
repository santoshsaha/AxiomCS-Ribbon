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
using System.Windows.Forms;



namespace AxiomIRISRibbon.SForceEdit
{
    /// <summary>
    /// Interaction logic for AttachmentsView.xaml
    ///  NEW File Added by PES
    /// </summary>
    public partial class AttachmentsView : RadWindow
    {

        private Data _d;
        DataTable _dt;


        private string _matterid;
        private string _versionid;
        private string _templateid;
        private string _versionName;
        private string _versionNumber;

        public AttachmentsView()
        {
            InitializeComponent();
            AxiomIRISRibbon.Utility.setTheme(this);

            _d = Globals.ThisAddIn.getData();
        }

        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }


        public void Create(DataTable dt, string versionid, string matterid, string templateid, string versionName, string versionNumber)
        {
            _dt = dt;
            //_matterid = matterid;
            //_versionid = versionid;
            //_templateid = templateid;
            //_versionName = versionName;
            //_versionNumber = versionNumber;



            //radComboAttachments.Items.Clear();

            //RadComboBoxItem i;
            //foreach (DataRow r in dt.Rows)
            //{


            //    if (r["Name"].ToString().Contains(".doc"))
            //    {
            //        i = new RadComboBoxItem();
            //        i.Tag = r["Id"].ToString();

            //        this.radComboAttachments.Items.Add(i);
            //        //dgFiles.Tag = r["Id"].ToString();
            //    }
            //}

            if (dgFiles.ItemsSource == null)
            {
                if (_dt.Rows.Count < 1) return;
                {
                    dgFiles.ItemsSource = _dt.DefaultView;
                }
            }

        }

        private void btnOpen_Click(object sender, RoutedEventArgs e)
        {

            if (this.dgFiles.SelectedItem == null)
            {
                System.Windows.Forms.MessageBox.Show("Select one document");
            }
            else
            {
                DataRowView dataRow = (DataRowView)dgFiles.SelectedItem;
                //int index = dgFiles.CurrentCell.Column.DisplayIndex;
                string attachmentid = dataRow.Row.ItemArray[0].ToString();
                //string attachmentid = ((RadComboBoxItem)(this.radComboAttachments.SelectedItem)).Tag.ToString();


                DataRow rw = _dt.AsEnumerable().Where(p => p.Field<string>("Id") == attachmentid).FirstOrDefault();
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
                csb.Create(filename, _versionid, _matterid, _templateid, _versionName, _versionNumber, attachmentid);

                elHost.Child = csb;
                elHost.Dock = System.Windows.Forms.DockStyle.Fill;
                System.Windows.Forms.UserControl u = new System.Windows.Forms.UserControl();
                u.Controls.Add(elHost);
                Microsoft.Office.Tools.CustomTaskPane taskPaneValue = Globals.ThisAddIn.CustomTaskPanes.Add(u, "Axiom IRIS Compare", doc.ActiveWindow);
                taskPaneValue.Visible = true;
                taskPaneValue.Width = 400;

                Globals.Ribbons.Ribbon1.CloseWindows();
                this.Close();
            }
        }

        //private void radComboAttachments_SelectionChanged(object sender, SelectionChangedEventArgs e)
        //{

        //    if (this.radComboAttachments.SelectedItem == null)
        //    {
        //        System.Windows.MessageBox.Show("Select one document");
        //    }
        //    else
        //    {

        //        string attachmentid = ((RadComboBoxItem)(this.radComboAttachments.SelectedItem)).Tag.ToString();

        //        DataRow rw = _dt.AsEnumerable().Where(p => p.Field<string>("Id") == attachmentid).FirstOrDefault();
        //        byte[] toBytes = Convert.FromBase64String(rw["body"].ToString());
        //        string filename = _d.GetTempFilePath(rw["Id"].ToString() + "_" + rw["Name"].ToString());

        //        File.WriteAllBytes(filename, toBytes);
        //        // _source = app.Documents.Open(filename);
        //        // rtbPreview.AppendText(filename);

        //        try
        //        {
        //            // Create an instance of StreamReader to read from a file. 
        //            // The using statement also closes the StreamReader. 
        //            using (StreamReader sr = new StreamReader(filename))
        //            {

        //                rtbPreview.AppendText(sr.ReadToEnd());

        //            }
        //        }
        //        catch (Exception exe)
        //        {
        //            //// Let the user know what went wrong. 
        //            //Console.WriteLine("The file could not be read:");
        //            //Console.WriteLine(exe.Message);
        //            System.Windows.MessageBox.Show(exe.ToString());
        //        }

        //        //rtbPreview.AppendText(filename);


        //    }

        //}


        private void dgFiles_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            //string attachmentid = dgFiles.Tag.ToString();
            //string attachmentid = dgFiles.SelectedValue.ToString();

            object nullobj = System.Reflection.Missing.Value;
            DataRowView dataRow = (DataRowView)dgFiles.SelectedItem;
            //int index = dgFiles.CurrentCell.Column.DisplayIndex;
            string attachmentid = dataRow.Row.ItemArray[0].ToString();

            DataRow rw = _dt.AsEnumerable().Where(p => p.Field<string>("Id") == attachmentid).FirstOrDefault();
            byte[] toBytes = Convert.FromBase64String(rw["body"].ToString());
            string filename = _d.GetTempFilePath(rw["Id"].ToString() + "_" + rw["Name"].ToString());
            File.WriteAllBytes(filename, toBytes);
            Word.Document doc = Globals.ThisAddIn.Application.Documents.Add(filename, ref nullobj, ref nullobj, false);
            //doc.TrackRevisions = true;
            rtbPreview.SelectAll();
            rtbPreview.Selection.Text = "";

            try
            {
                if (doc.ActiveWindow.Selection.StoryLength > 0)
                {
                    doc.ActiveWindow.Selection.WholeStory();

                    doc.ActiveWindow.Selection.Copy();
                    rtbPreview.Paste();
                    rtbPreview.IsReadOnly = true;
                    doc.Close();
                }
                this.Focus();
            }
            catch (Exception exe)
            {

                System.Windows.Forms.MessageBox.Show(exe.ToString());
            }
            finally
            {
                // doc.Close();
            }

        }
        private void FilesRowDoubleClick(object sender, RoutedEventArgs e)
        {
            DataRowView dataRow = (DataRowView)dgFiles.SelectedItem;
            //int index = dgFiles.CurrentCell.Column.DisplayIndex;
            string attachmentid = dataRow.Row.ItemArray[0].ToString();
            DataRow rw = _dt.AsEnumerable().Where(p => p.Field<string>("Id") == attachmentid).FirstOrDefault();
            byte[] toBytes = Convert.FromBase64String(rw["body"].ToString());
            string filename = _d.GetTempFilePath(rw["Id"].ToString() + "_" + rw["Name"].ToString());

            File.WriteAllBytes(filename, toBytes);
            // _source = app.Documents.Open(filename);


            Word.Document doc = Globals.ThisAddIn.Application.Documents.Add(filename);

            System.Windows.Forms.Integration.ElementHost elHost = new System.Windows.Forms.Integration.ElementHost();
            SForceEdit.CompareSideBar csb = new SForceEdit.CompareSideBar();
            csb.Create(filename, _versionid, _matterid, _templateid, _versionName, _versionNumber, attachmentid);

            elHost.Child = csb;
            elHost.Dock = System.Windows.Forms.DockStyle.Fill;
            System.Windows.Forms.UserControl u = new System.Windows.Forms.UserControl();
            u.Controls.Add(elHost);
            Microsoft.Office.Tools.CustomTaskPane taskPaneValue = Globals.ThisAddIn.CustomTaskPanes.Add(u, "Axiom IRIS Compare", doc.ActiveWindow);
            taskPaneValue.Visible = true;
            taskPaneValue.Width = 400;

            Globals.Ribbons.Ribbon1.CloseWindows();
            this.Close();

        }

        private void radComboAttachments_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

        }
    }
}