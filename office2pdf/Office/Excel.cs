using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;

namespace office2pdf.Office
{
    public class Excel : Office
    {
        public const string EXT = @".xlsx";

        public Excel()
        {
            this.ext = EXT;
        }

        /// <summary>
        /// Excelファイルかどうかをチェック
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        new public static bool IsValid(string path)
        {
            return Office.MatchExt(path, string.Format(@".*\{0}$", EXT));
        }

        /// <summary>
        /// ファイルをPDF形式で保存
        /// </summary>
        override public void SavePdf()
        {
            Microsoft.Office.Interop.Excel.Application app = null;
            Microsoft.Office.Interop.Excel.Workbooks books = null;
            Microsoft.Office.Interop.Excel.Workbook book = null;

            try
            {
                app = new Microsoft.Office.Interop.Excel.Application();
                books = app.Workbooks;

                book = books.Open(this.GetAbsolutePath());

                book.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, this.GetPdfPath(), XlFixedFormatQuality.xlQualityStandard);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                Console.WriteLine(e.StackTrace);
            }
            finally
            {
                if (book != null)
                {
                    book.Close();
                }
                if (app != null)
                {
                    app.Quit();
                }
            }
        }
    }
}
