﻿using System;
using System.IO;
using ClosedXML.Excel;
using System.Net;
using System.Linq;

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
            string[] 英単語 = ex.読み取り();
            Web操作 we = new Web操作(英単語);
            string[] 発音記号 = we.読み取り();
            ex.書き込み(発音記号);
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
            lastRow = worksheet.LastRowUsed().RowNumber();
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
        public void 書き込み(string[] 発音記号)
        {
            for (int i = 0; i < 発音記号.Length; i++)
            {
                worksheet.Cell(i + 1, 2).SetValue(発音記号[i]);
            }
            workbook.Save();
        }
    }

    class Web操作
    {
        private int count;
        private string[] url = new string[0];
        const string weblio_url = @"https://ejje.weblio.jp/content/";
        private string[] html = new string[0];
        public Web操作(string[] 英単語)
        {
            count = 英単語.Length;
            Array.Resize<string>(ref url, count);
            Array.Resize<string>(ref html, count);
            url = (Array.ConvertAll(英単語, conecturl));
        }

        private string conecturl(string input) { return $"{weblio_url}{input}"; }

        public string[] 読み取り()
        {
            WebClient wc = new WebClient();
            string[] 読み方 = new string[count];
            var sw = new System.Diagnostics.Stopwatch();
            for (int i = 0; i < count; i++)
            {
                sw.Restart();
                html[i] = wc.DownloadString(url[i]);
                sw.Stop();
                Console.WriteLine(sw.Elapsed);
                string[] 分割 = html[i].Split("\n");
                int 検索行 = 0;
                int 符号 = 1;
                int c = 1;
                while (true)
                {
                    int 行 = 分割[検索行].IndexOf(@"</span><span class=phoneticEjjeDc>(米国英語)");
                    if (行 != -1)
                    {
                        読み方[i] = 分割[検索行].Substring(0, 行);
                        読み方[i] = 読み方[i].Substring(92);
                        // Console.WriteLine(読み方[i]);
                        break;
                    }
                    検索行 += 1;
                    符号 *= -1;
                    c++;
                }
            }
            return 読み方;
        }
    }
}