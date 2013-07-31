using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;

namespace ModelUN
{
    partial class About : Form
    {
        private readonly string TITLE = "About Model UN Seat Assign";

        public About()
        {
            InitializeComponent();
            this.Text = TITLE;
        }

        private void labelCompanyName_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("http://alex-vallejo.com");
        }

        private void okButton_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("http://alex-vallejo.com");
        }
    }
}
