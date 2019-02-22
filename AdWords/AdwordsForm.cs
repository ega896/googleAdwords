﻿using System;
using System.IO;
using System.Windows.Forms;

namespace AdWords
{
    public partial class AdWordsForm : Form
    {
        public AdWordsForm()
        {
            InitializeComponent();
        }

        private void chooseFileButton_Click(object sender, EventArgs e)
        {
            var result = openFileDialog.ShowDialog();
            var extension = Path.GetExtension(openFileDialog.FileName);
            if (result == DialogResult.OK)
            {
                if (extension != ".xlsx" && extension != ".xls")
                {
                    MessageBox.Show(@"Choose file with ""/"".xlsx""/"" or ""/.xls"" format");
                    return;
                }

                filePathLabel.Text = openFileDialog.FileName;
                executeButton.Visible = true;
                progressBar.Value = 0;
                doneLabel.Text = null;
            }
            else
            {
                MessageBox.Show(@"Unexpected error. Please try again");
            }
        }

        private void executeButton_Click(object sender, EventArgs e)
        {
            try
            {
                var excelService = new ExcelService(filePathLabel.Text);
                doneLabel.Text = excelService.Execute(progressBar);
            }
            catch (Exception)
            {
                MessageBox.Show(@"Unexpected error. Please try again");
            }
        }
    }
}
