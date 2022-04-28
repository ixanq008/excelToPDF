using Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;

namespace ExcelToPDF2
{
    public class ExcelToPDF2
    {
        /// <summary>
        /// excel转换PDF
        /// </summary>
        /// <param name="sourceFile">原文件 如："D:\配方关联导入模板.xlsx"</param>
        /// <param name="targetFile">目标文件 如："D:\配方关联导入模板.pdf"</param>
        /// <returns></returns>
        public static bool XLSConvertToPDF(string sourceFile,string targetFile)
        {
            string sourcePath = sourceFile;
            string targetPath = targetFile;
            bool result = false;
            Excel.XlFixedFormatType targetType = Microsoft.Office.Interop.Excel.XlFixedFormatType.xlTypePDF;// Excel.XlFixedFormatType.xlTypePDF;
            object paramMissing = Type.Missing;
            Excel.ApplicationClass application = null;
            Excel.Workbook excelWorkBook = null;
            try
            {
                application = new Excel.ApplicationClass();
                object target = targetPath;
                object type = targetType;
                excelWorkBook = application.Workbooks.Open(sourcePath);

                Excel.Worksheet xlSheet = (Excel.Worksheet)excelWorkBook.Worksheets[1];

                if (excelWorkBook != null)
                {
                    excelWorkBook.ExportAsFixedFormat(targetType, target, Excel.XlFixedFormatQuality.xlQualityStandard, true, true, paramMissing, paramMissing, false, paramMissing);
                    result = true;
                }
            }
            catch
            {
                result = false;
            }
            finally
            {
                if (excelWorkBook != null)
                {
                    excelWorkBook.Close(false, paramMissing, paramMissing);
                    excelWorkBook = null;
                }
                if (application != null)
                {
                    application.Quit();
                    application = null;
                }

                //ms solution is like this
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            return result;
        }

        private static string GetFileName(string pathString)
        {
            return Path.ChangeExtension(pathString, @".pdf");
        }
    }
}
