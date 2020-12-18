using System;
using System.IO;
using ClosedXML.Excel;
namespace 発音記号
{
    class 発音記号
    {
        static void Main(string[] args)
        {
            string exeファイルのパス = AppDomain.CurrentDomain.BaseDirectory.TrimEnd('\\');
            DirectoryInfo 作業ディレクトリ = new DirectoryInfo(exeファイルのパス);
            さかのぼる(ref 作業ディレクトリ, 3);
            Excel操作 ex = new Excel操作(作業ディレクトリ);
            int length = ex.読み取り();
            Console.WriteLine(length);
        }

        static void さかのぼる(ref DirectoryInfo path, int 回数)
        {
            for (int i = 0; i < 回数; i++)
            {
                path = path.Parent;
            }
        }
    }

    class Excel操作
    {
        private string excelのパス;
        private XLWorkbook workbook;
        private IXLWorksheet worksheet;
        int lastRow;
        public Excel操作(DirectoryInfo 作業ディレクトリ)
        {
            excelのパス = 作業ディレクトリ.ToString() + @"\発音記号.xlsx";
            workbook = new XLWorkbook(excelのパス);
            worksheet = workbook.Worksheet("Sheet1");
        }
        public int 読み取り()
        {
            lastRow = worksheet.LastRowUsed().RowNumber();
            return lastRow;
        }
    }

    class Web操作
    {
        public Web操作()
        {

        }
    }
}