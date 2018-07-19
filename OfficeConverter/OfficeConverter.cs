using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Core;

namespace OfficeConverterLib
{
    public static class OfficeConverter
    {
        public static void ConvertWordDocumentToPdf(string srcFile, string dstFile)
        {
            Word.Application wordApp = null;
            Word.Document wordDoc = null;
            try
            {
                wordApp = new Word.Application
                {
                    DisplayAlerts = Word.WdAlertLevel.wdAlertsNone
                };
                wordDoc = wordApp.Documents.OpenNoRepairDialog(srcFile, true, false, false);
                wordDoc.ExportAsFixedFormat(dstFile, Word.WdExportFormat.wdExportFormatPDF);
            }
            finally
            {
                wordDoc?.Close(SaveChanges: false);
                wordApp?.Quit();
            }
        }

        public static void ConvertExcelDocumentToPdf(string srcFile, string dstFile)
        {
            Excel.Application excelApp = null;
            Excel.Workbook excelWorkBook = null;
            try
            {
                excelApp = new Excel.Application();
                excelWorkBook = excelApp.Workbooks.Open(
                    srcFile,
                    ReadOnly: true,
                    IgnoreReadOnlyRecommended: true,
                    Notify: false
                );
                excelWorkBook.ExportAsFixedFormat(
                    Type: Excel.XlFixedFormatType.xlTypePDF,
                    Filename: dstFile,
                    Quality: Excel.XlFixedFormatQuality.xlQualityStandard,
                    IncludeDocProperties: true,
                    OpenAfterPublish: false,
                    IgnorePrintAreas: true
                );
            }
            finally
            {
                excelWorkBook?.Close(SaveChanges: false);
                excelApp?.Quit();
            }
        }

        public static void ConvertPowerPointDocumentToPdf(string srcFile, string dstFile)
        {
            PowerPoint.Application powerPointApp = null;
            PowerPoint.Presentation powerPointPresentation = null;
            try
            {
                powerPointApp = new PowerPoint.Application();
                powerPointPresentation = powerPointApp.Presentations.Open(srcFile,
                    MsoTriState.msoTrue,
                    MsoTriState.msoFalse,
                    MsoTriState.msoFalse);
                powerPointPresentation.ExportAsFixedFormat(dstFile, PowerPoint.PpFixedFormatType.ppFixedFormatTypePDF);
            }
            finally
            {
                powerPointPresentation?.Close();
                powerPointApp?.Quit();
            }
        }
    }
}
