using System;
using System.Windows.Forms;
using System.IO;
using System.Diagnostics;

namespace OfficeInteropClient
{
    public partial class OfficeConverter : Form
    {
        private string wordFilter = "Word|*.docx";
        private string excelFilter = "Excel|*.xlsx";
        private string powerFilter = "Word|*.ppt";

        public OfficeConverter()
        {
            InitializeComponent();
        }

        private void btnWord_Click(object sender, EventArgs e)
        {
            string srcFile = ShowFileDialog(wordFilter);
            if (String.IsNullOrEmpty(srcFile)) return;
            string dstFile = GetResultFileName(srcFile);

            OfficeConverterLib.OfficeConverter.ConvertWordDocumentToPdf(srcFile, dstFile);
            Debug.WriteLine($"[OK]: Exported ({dstFile})");
        }

        private void btnExcel_Click(object sender, EventArgs e)
        {
            try
            {
                string srcFile = ShowFileDialog(excelFilter);
                if (String.IsNullOrEmpty(srcFile)) return;
                string dstFile = GetResultFileName(srcFile);

                OfficeConverterLib.OfficeConverter.ConvertExcelDocumentToPdf(srcFile, dstFile);
                Debug.WriteLine($"[OK]: Exported ({dstFile})");
            }
            catch (Exception ex)
            {
                Debug.WriteLine("[ERROR]: " + ex.Message);
            }
        }

        private void btnPowerPoint_Click(object sender, EventArgs e)
        {
            string srcFile = ShowFileDialog(powerFilter);
            if (String.IsNullOrEmpty(srcFile)) return;
            string dstFile = GetResultFileName(srcFile);

            OfficeConverterLib.OfficeConverter.ConvertPowerPointDocumentToPdf(srcFile, dstFile);
            Debug.WriteLine($"[OK]: Exported ({dstFile})");
        }

        private string ShowFileDialog(string filter)
        {
            if (String.IsNullOrEmpty(filter))
            {
                throw new ArgumentNullException();
            }
            System.Windows.Forms.FileDialog dialog = new OpenFileDialog()
            {
                Filter = filter,
                InitialDirectory = @"C:",
            };

            if (dialog.ShowDialog() == DialogResult.OK)
            {
                return dialog.FileName;
            }

            return null;
        }

        private string GetResultFileName(string fileName)
        {
            string resFileName = Path.GetFileName(Path.ChangeExtension(fileName, "pdf"));
            string docPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            
            return docPath + Path.DirectorySeparatorChar + resFileName;
        }
    }
}
