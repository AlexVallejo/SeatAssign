using System;
using System.Diagnostics;
using System.Windows.Forms;

namespace ModelUN
{
    partial class About : Form
    {

        public About()
        {
            InitializeComponent();
        }

        private void close_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            ProcessStartInfo sInfo = new ProcessStartInfo("http://alex-vallejo.com/");
            Process.Start(sInfo);
        }
    }
}