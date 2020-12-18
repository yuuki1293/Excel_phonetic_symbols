using System;
using System.IO;
using ClosedXML.Excel;
using System.Net;

namespace 発音記号
{
    class 発音記号
    {
        static void Main(string[] args)
        {
            string exeファイルのパス = AppDomain.CurrentDomain.BaseDirectory.TrimEnd('\\');
            DirectoryInfo 作業ディレクトリ = new DirectoryInfo(exeファイルのパス);
            さかのぼる(ref 作業ディレクトリ, 3);
            int length = 0;
            Excel操作 ex = new Excel操作(作業ディレクトリ, ref length);
            string[] 英単語 = new string[length];
            英単語 = ex.読み取り();
            Web操作 we = new Web操作(英単語, length);
            string[] 発音記号 = new string[length];
            発音記号 = we.読み取り();
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
        public Excel操作(DirectoryInfo 作業ディレクトリ, ref int length)
        {
            excelのパス = 作業ディレクトリ.ToString() + @"\発音記号.xlsx";
            workbook = new XLWorkbook(excelのパス);
            worksheet = workbook.Worksheet("Sheet1");
            length = worksheet.LastRowUsed().RowNumber();
            lastRow = length;
        }
        public string[] 読み取り()
        {
            string[] 英単語 = new string[lastRow];
            for (int i = 0; i < lastRow; i++)
            {
                英単語[i] = worksheet.Cell(i + 1, 1).Value.ToString();
            }
            return 英単語;
        }
    }

    class Web操作
    {
        string reg = @"<span class=phoneticEjjeDesc>həlóʊ.+?<//span><span class=phoneticEjjeDc>(米国英語)";
        private int count;
        private string[] url = new string[0];
        const string weblio_url = @"https://ejje.weblio.jp/content/";
        private string[] html = new string[0];
        public Web操作(string[] 英単語, int length)
        {
            count = length;
            Array.Resize<string>(ref url, length);
            Array.Resize<string>(ref html, length);
            for (int i = 0; i < length; i++)
            {
                url[i] = weblio_url + 英単語[i];
            }
        }
        public string[] 読み取り()
        {
            WebClient wc = new WebClient();
            for (int i = 0; i < count; i++)
            {
                html[i] = wc.DownloadString(url[i]);
            }

            return html;
        }
    }
}