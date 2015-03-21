using System;
using System.IO;
using System.Reflection;

namespace TdqqClient.Services.Export
{
    /// <summary>
    /// 通过调用Office组件的将Excel文件转换成pdf文件
    /// </summary>
    class Export2Pdf
    {
        /// <summary>
        /// 将Excel转成PDF文件
        /// </summary>
        /// <param name="sourceFile">Excel文件路径</param>
        /// <param name="targetFile">Pdf文件路径</param>
        public static void Excel2Pdf(string sourceFile, string targetFile)
        {

            if (!File.Exists(sourceFile)) return;
            object objOpt = Missing.Value;

            Microsoft.Office.Interop.Excel.Application excelApp = null;
            try
            {
                excelApp = new Microsoft.Office.Interop.Excel.Application();
                excelApp.Workbooks.Open(sourceFile, objOpt, objOpt, objOpt, objOpt, objOpt, true, objOpt, objOpt, true, objOpt, objOpt, objOpt, objOpt, objOpt);
                excelApp.ActiveWorkbook.ExportAsFixedFormat(Microsoft.Office.Interop.Excel.XlFixedFormatType.xlTypePDF, (object)targetFile, objOpt, objOpt, objOpt, objOpt, objOpt, objOpt, objOpt);
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                if (excelApp != null)
                    excelApp.Quit();
            }
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
        /// <summary>
        /// 将Excel转成PDF文件(同名)
        /// </summary>
        /// <param name="sourceFile">Excel文件路径</param>
        public static void Excel2Pdf(string sourceFile)
        {

            if (!File.Exists(sourceFile)) return;
            object objOpt = Missing.Value;
            
            var targetFile = sourceFile.Substring(0, sourceFile.LastIndexOf(".")) + @".pdf";
            Microsoft.Office.Interop.Excel.Application excelApp = null;
            try
            {
                excelApp = new Microsoft.Office.Interop.Excel.Application();
                excelApp.Workbooks.Open(sourceFile, objOpt, objOpt, objOpt, objOpt, objOpt, true, objOpt, objOpt, true, objOpt, objOpt, objOpt, objOpt, objOpt);
                excelApp.ActiveWorkbook.ExportAsFixedFormat(Microsoft.Office.Interop.Excel.XlFixedFormatType.xlTypePDF, (object)targetFile, objOpt, objOpt, objOpt, objOpt, objOpt, objOpt, objOpt);
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.ToString());
            }
            finally
            {
                if (excelApp != null)
                    excelApp.Quit();
            }
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }
}
