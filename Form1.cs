using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace WordPreetiToUnicode
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            InitializeCustomUI();
        }

        private void InitializeCustomUI()
        {
            this.Text = "Preeti to Unicode Converter";
            this.Size = new System.Drawing.Size(300, 150);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.TopMost = true;

            if (File.Exists("app.ico"))
            {
                this.Icon = new System.Drawing.Icon("app.ico");
            }

            Button convertButton = new Button();
            convertButton.Text = "Convert Selected Text";
            convertButton.Size = new System.Drawing.Size(200, 50);
            convertButton.Location = new System.Drawing.Point(40, 30);
            convertButton.Click += ConvertButton_Click;

            this.Controls.Add(convertButton);
        }

        private void ConvertButton_Click(object? sender, EventArgs e)
        {
            dynamic? wordApp = null;
            dynamic? currentSelection = null;

            try
            {
                object comObj = ComSupport.GetActiveObject("Word.Application");
                wordApp = comObj;
                currentSelection = wordApp.Selection;

                string rawPreetiText = currentSelection.Text;

                if (string.IsNullOrEmpty(rawPreetiText) || rawPreetiText == "\r")
                {
                    MessageBox.Show("Please select some Preeti text inside MS Word first.", "No Selection", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                string unicodeText = PreetiEngine.ConvertBlock(rawPreetiText);

                currentSelection.Text = unicodeText;
                currentSelection.Font.Name = "Kalimati";
            }
            catch (COMException)
            {
                MessageBox.Show("Could not connect to MS Word. Make sure MS Word is open.", "Connection Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (currentSelection != null) Marshal.ReleaseComObject(currentSelection);
                if (wordApp != null) Marshal.ReleaseComObject(wordApp);
            }
        }
    }
}