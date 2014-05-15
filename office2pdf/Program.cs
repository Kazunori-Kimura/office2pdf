using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using office2pdf.Office;

namespace office2pdf
{
    class Program
    {
        static void Main(string[] args)
        {
            foreach (string path in args)
            {
                ToPdf(path);
            }
        }

        static void ToPdf(string path)
        {
            Office.Office office;

            if (Excel.IsValid(path))
            {
                office = new Excel();
            }
            else if (PowerPoint.IsValid(path))
            {
                office = new PowerPoint();
            }
            else if (Word.IsValid(path))
            {
                office = new Word();
            }
            else
            {
                return;
            }

            Console.Write(string.Format(@"{0} -> ", path));
            office.path = path;

            office.SavePdf();
        }
    }
}
