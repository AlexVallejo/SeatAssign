using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace ModelUN
{
    public partial class ModelUNAssign : Form
    {
        public ModelUNAssign()
        {
            InitializeComponent();
        }

        private void ModelUNForm_Load(object sender, EventArgs e)
        {
            //Implement drag and drop capibilities here
            //this.AllowDrop = true;
        }

        private void process_Click(object sender, EventArgs e)
        {
            progressBar.Minimum = 0;
            progressBar.Maximum = 50;
            progressBar.Value = progressBar.Minimum;
            progressBar.Step = 1;

            for (int i = 0; i <= 50; i++)
            {
                System.Threading.Thread.Sleep(100);
                progressBar.PerformStep();
            }

            MessageBox.Show("A worksheet has been added to the original spreadsheet with the seat assignments.");
        }

        private void inputSelect_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Filter = "Microsoft Excel files |*.xlsx";
            dialog.Title = "Select the Excel file with the applicant information";

            if (dialog.ShowDialog() == DialogResult.OK)
            {
                string file = dialog.FileName;
                Console.WriteLine(file);

                try
                {
                    //read file here
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.ToString());
                }
            }
            
        }

        private void aboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            About form = new About();
            form.Show();
        }

        private void instructionsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Instructions command executed");
        }
    }
}