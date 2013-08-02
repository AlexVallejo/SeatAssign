using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Core;
using System.IO;

namespace ModelUN
{
    public partial class ModelUNAssign : Form
    {
        private string filepath;

        public ModelUNAssign()
        {
            InitializeComponent();
        }

        private void ModelUNForm_Load(object sender, EventArgs e)
        {
            //Implement drag and drop capibilities here
            //this.AllowDrop = true;

            //Initialize the progress bar
            progressBar.Minimum = 0;
            progressBar.Maximum = 100;
            progressBar.Value = progressBar.Minimum;
            progressBar.Step = 15;
        }

        private void process_Click(object sender, EventArgs e)
        {
            if (!validRequisites())
                return;

            new ApplicantAssign(filepath, progressBar);

            MessageBox.Show("A worksheet has been added to the original spreadsheet with the seat assignments.");
        }

        private bool validRequisites()
        {
            //verify Excel is closed
            if (System.Diagnostics.Process.GetProcessesByName("excel").Length > 0)
            {
                MessageBox.Show("You must close Microsoft Excel before using this tool.");
                return false;
            }

            if (filepath == null)
            {
                MessageBox.Show("You must select a valid input file first.");
                return false;
            }

            if (!File.Exists("countries.txt"))
            {
                MessageBox.Show("Ensure the 'countries.txt' file exists");
                return false;
            }

            if (!File.Exists("regions.txt"))
            {
                MessageBox.Show("Ensure the 'regions.txt' file exists");
                return false;
            }
            
            return true;
         }

        private void inputSelect_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Filter = "Microsoft Excel files |*.xlsx";
            dialog.Title = "Select the Excel file with the applicant information";

            if (dialog.ShowDialog() == DialogResult.OK)
            {
                filepath = dialog.FileName;
                Console.WriteLine(filepath);

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
            Instructions form = new Instructions();
            form.Show();
        }
    }
}