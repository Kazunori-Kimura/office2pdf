using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;


namespace office2pdf.Office
{
    class Word : Office
    {
        public const string EXT = @".docx";

        public Word()
        {
            this.ext = EXT;
        }

        /// <summary>
        /// Wordファイルかどうかをチェック
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
            Microsoft.Office.Interop.Word.Application word = null;
            Microsoft.Office.Interop.Word.Documents docs = null;
            Microsoft.Office.Interop.Word.Document d = null;

            try
            {
                //ファイルを取得
                object file = this.GetAbsolutePath();
                word = new Microsoft.Office.Interop.Word.Application();
                docs = word.Documents;
                d = docs.Open(file);

                d.ExportAsFixedFormat(this.GetPdfPath(),
                    Microsoft.Office.Interop.Word.WdExportFormat.wdExportFormatPDF,
                    false,
                    Microsoft.Office.Interop.Word.WdExportOptimizeFor.wdExportOptimizeForPrint,
                    Microsoft.Office.Interop.Word.WdExportRange.wdExportAllDocument);

            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                Console.WriteLine(e.StackTrace);
            }
            finally
            {
                if (d != null)
                {
                    d.Close();
                }
                if (word != null)
                {
                    word.Quit();
                }
            }
        }

    }
}
