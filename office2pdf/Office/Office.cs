using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace office2pdf.Office
{
    public abstract class Office
    {
        public string path { get; set; }

        protected string ext;

        /// <summary>
        /// 絶対パスに変換する
        /// </summary>
        /// <returns></returns>
        protected string GetAbsolutePath()
        {
            if (!System.IO.File.Exists(this.path))
            {
                throw new System.IO.IOException("指定されたファイルが存在しません。");
            }

            return System.IO.Path.GetFullPath(this.path);
        }

        /// <summary>
        /// PDFファイルパスを取得する
        /// </summary>
        /// <returns></returns>
        protected string GetPdfPath()
        {
            return this.GetPdfPath(this.path);
        }

        /// <summary>
        /// PDFファイルパスを取得する
        /// </summary>
        /// <returns>path</returns>
        protected string GetPdfPath(string filePath)
        {
            string pdfExt = @".pdf";
            string pattern = string.Format(@"\{0}$", this.ext);
            var reg = new Regex(pattern);

            //置換
            string pdfPath = reg.Replace(filePath, pdfExt);
            Console.WriteLine(pdfPath);
            return pdfPath;
        }

        /// <summary>
        /// PDFファイルを保存する
        /// </summary>
        public abstract void SavePdf();

        /// <summary>
        /// ファイルの拡張子から、Word/Excel/PowerPointを判断する
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        public static bool IsValid(string filePath)
        {
            return false;
        }

        /// <summary>
        /// 拡張子が一致するかをチェックする
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="ext"></param>
        /// <returns></returns>
        protected static bool MatchExt(string filePath, string ext)
        {
            string pattern = string.Format(@"\.{0}$", ext);
            return Regex.IsMatch(filePath, pattern, RegexOptions.IgnoreCase);
        }
    }
}
