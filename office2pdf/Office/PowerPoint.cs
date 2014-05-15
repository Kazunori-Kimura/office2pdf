using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.PowerPoint;

namespace office2pdf.Office
{
    public class PowerPoint : Office
    {
        public const string EXT = @".pptx";

        public PowerPoint()
        {
            this.ext = EXT;
        }

        /// <summary>
        /// PowerPointファイルかどうかをチェック
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
            //ファイルを取得
            string file = this.GetAbsolutePath();

            Microsoft.Office.Interop.PowerPoint.Application ppt = null;
            Microsoft.Office.Interop.PowerPoint.Presentation p = null;

            try
            {
                ppt = new Microsoft.Office.Interop.PowerPoint.Application();

                //ファイルを開く
                p = ppt.Presentations.Open(file);

                //PDFとして保存
                p.SaveAs(this.GetPdfPath(),
                    PpSaveAsFileType.ppSaveAsPDF,
                    Microsoft.Office.Core.MsoTriState.msoTrue);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                Console.WriteLine(e.StackTrace);
            }
            finally
            {
                if(p != null)
                {
                    p.Close();
                }
                if (ppt != null)
                {
                    ppt.Quit();
                }
            }
        }

    }
}
