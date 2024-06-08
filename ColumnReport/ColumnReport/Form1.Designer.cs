namespace ColumnReport
{
    partial class MainForm
    {
        private System.ComponentModel.IContainer components = null;
        private System.Windows.Forms.TextBox pathTextBox;
        private System.Windows.Forms.Button browseButton;
        private System.Windows.Forms.Button exportButton;
        private System.Windows.Forms.Button closeButton;
        private System.Windows.Forms.Button linkedInButton;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        private void InitializeComponent()
        {
            this.pathTextBox = new System.Windows.Forms.TextBox();
            this.browseButton = new System.Windows.Forms.Button();
            this.exportButton = new System.Windows.Forms.Button();
            this.closeButton = new System.Windows.Forms.Button();
            this.linkedInButton = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // pathTextBox
            // 
            this.pathTextBox.Location = new System.Drawing.Point(13, 13);
            this.pathTextBox.Name = "pathTextBox";
            this.pathTextBox.Size = new System.Drawing.Size(200, 20);
            this.pathTextBox.TabIndex = 0;
            // 
            // browseButton
            // 
            this.browseButton.Location = new System.Drawing.Point(220, 10);
            this.browseButton.Name = "browseButton";
            this.browseButton.Size = new System.Drawing.Size(75, 23);
            this.browseButton.TabIndex = 1;
            this.browseButton.Text = "Browse";
            this.browseButton.UseVisualStyleBackColor = true;
            this.browseButton.Click += new System.EventHandler(this.browseButton_Click);
            // 
            // exportButton
            // 
            this.exportButton.Location = new System.Drawing.Point(13, 50);
            this.exportButton.Name = "exportButton";
            this.exportButton.Size = new System.Drawing.Size(75, 23);
            this.exportButton.TabIndex = 2;
            this.exportButton.Text = "Export";
            this.exportButton.UseVisualStyleBackColor = true;
            this.exportButton.Click += new System.EventHandler(this.exportButton_Click);
            // 
            // closeButton
            // 
            this.closeButton.Location = new System.Drawing.Point(115, 50);
            this.closeButton.Name = "closeButton";
            this.closeButton.Size = new System.Drawing.Size(75, 23);
            this.closeButton.TabIndex = 3;
            this.closeButton.Text = "Close";
            this.closeButton.UseVisualStyleBackColor = true;
            this.closeButton.Click += new System.EventHandler(this.closeButton_Click);
            // 
            // linkedInButton
            // 
            this.linkedInButton.Location = new System.Drawing.Point(220, 50);
            this.linkedInButton.Name = "linkedInButton";
            this.linkedInButton.Size = new System.Drawing.Size(75, 23);
            this.linkedInButton.TabIndex = 4;
            this.linkedInButton.Text = "LinkedIn";
            this.linkedInButton.UseVisualStyleBackColor = true;
            this.linkedInButton.Click += new System.EventHandler(this.linkedInButton_Click);
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(304, 91);
            this.Controls.Add(this.linkedInButton);
            this.Controls.Add(this.closeButton);
            this.Controls.Add(this.exportButton);
            this.Controls.Add(this.browseButton);
            this.Controls.Add(this.pathTextBox);
            this.Name = "MainForm";
            this.Text = "Column Report Exporter";
            this.ResumeLayout(false);
            this.PerformLayout();
        }
    }
}
