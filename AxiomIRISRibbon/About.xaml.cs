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
using System.Configuration;
using System.Deployment;
using Microsoft.Win32;
using System.Diagnostics;

namespace AxiomIRISRibbon
{
    /// <summary>
    /// Interaction logic for UserControl1.xaml
    /// </summary>
    public partial class About : RadWindow
    {
        public About()
        {
            InitializeComponent();
            Utility.setTheme(this);

            string version = GetRunningVersion().ToString();
            this.tbVersion.Text = version.ToString();

            this.tbSF.Text = Globals.ThisAddIn.getData().GetInstanceInfo();
            this.tbUser.Text = Globals.ThisAddIn.getData().GetUserInfo();

            //Code PES
            string timestamp = RetrieveLinkerTimestamp().ToString();
            if (timestamp == string.Empty)
            { timestamp = "UNKOWN!"; }
            this.verTimestamp.Text = "Build Time : " + timestamp;
            //End Code
        }

        private void btnOK_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private string GetRunningVersion()
        {
            // TO DO
            string v = "IRIS Ribbon | Version - UNKOWN!";
          //  string v = "IRIS Ribbon | Version 1.00.00.05";
            try
            {
                    System.Deployment.Application.ApplicationDeployment ad = System.Deployment.Application.ApplicationDeployment.CurrentDeployment;
                    Version vrn = ad.CurrentVersion;
                    v = "IRIS Ribbon | Version " + vrn.Major + "." + vrn.Minor + "." + vrn.Build + "." + vrn.Revision;
                
            }
            catch (Exception)
            {

            }

            return v;
        }
        //Code PES
        private DateTime RetrieveLinkerTimestamp()
        {
            string filePath = System.Reflection.Assembly.GetCallingAssembly().Location;
            const int c_PeHeaderOffset = 60;
            const int c_LinkerTimestampOffset = 8;
            byte[] b = new byte[2048];
            System.IO.Stream s = null;

            try
            {
                s = new System.IO.FileStream(filePath, System.IO.FileMode.Open, System.IO.FileAccess.Read);
                s.Read(b, 0, 2048);
            }
            finally
            {
                if (s != null)
                {
                    s.Close();
                }
            }

            int i = System.BitConverter.ToInt32(b, c_PeHeaderOffset);
            int secondsSince1970 = System.BitConverter.ToInt32(b, i + c_LinkerTimestampOffset);
            DateTime dt = new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc);
            dt = dt.AddSeconds(secondsSince1970);
            dt = dt.ToLocalTime();
            return dt;
        }
        // end code
        private void Hyperlink_Click(object sender, RoutedEventArgs e)
        {
            Globals.ThisAddIn.OpenAboutReleaseNotes();                    
        }

        private void windowAbout_Activated(object sender, EventArgs e)
        {
            this.tbSF.Text = Globals.ThisAddIn.getData().GetInstanceInfo();
            this.tbUser.Text = Globals.ThisAddIn.getData().GetUserInfo();
        }

    }
}
