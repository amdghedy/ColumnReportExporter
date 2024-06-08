using System;
using System.IO;
using System.Windows.Forms;

namespace ColumnReport
{
    public partial class MainForm : Form
    {
        public string FilePath { get; private set; }

        public MainForm()
        {
            InitializeComponent();
        }

        private void browseButton_Click(object sender, EventArgs e)
        {
            using (FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog())
            {
                if (folderBrowserDialog.ShowDialog() == DialogResult.OK)
                {
                    pathTextBox.Text = folderBrowserDialog.SelectedPath;
                }
            }
        }

        private void exportButton_Click(object sender, EventArgs e)
        {
            FilePath = Path.Combine(pathTextBox.Text, $"ColumnsReport-{DateTime.Now:yyyy-MM-dd_HH-mm-ss}.xlsx");
            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void closeButton_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }

        private void linkedInButton_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("https://www.linkedin.com/in/abdelrahim-deghidy-26b06538/");
        }
    }
}
