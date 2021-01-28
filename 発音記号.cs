using Microsoft.WindowsAPICodePack.Dialogs;
using System;
using System.IO;
using ClosedXML.Excel;
using System.Net;
using System.Text.RegularExpressions;
using System.Diagnostics;
using Microsoft.Win32;
using System.Text;

namespace 発音記号
{
    class 発音記号
    {
        static void Main(string[] args)
        {
            string exeファイルのパス = AppDomain.CurrentDomain.BaseDirectory.TrimEnd('\\');
            DirectoryInfo 作業ディレクトリ = new(exeファイルのパス);
            さかのぼる(ref 作業ディレクトリ, 5);
            Excel操作 ex = new(作業ディレクトリ);
            string[] 英単語 = ex.読み取り();
            Web操作 we = new(英単語);
            string[,] 発音記号 = we.読み取り();
            ex.書き込み(発音記号);
            Error.Massage("システムは正常に終了しました");
            ex.起動();
        }

        static int さかのぼる(ref DirectoryInfo path, int 回数)
        {
            for (int i = 0; i < 回数; i++)
            {
                path = path.Parent;
                if (File.Exists($"{path}/発音記号.xlsx")) { return 0; }
                // Console.WriteLine(i);
            }
            Error.Massage("発音記号.xlsxが見つかりません");
            Error.Exit();
            return -1;
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
            try
            {
                workbook = new XLWorkbook(excelのパス);
            }
            catch (System.IO.IOException)
            {
                Error.Massage("Excelが開いたままです。Excelを終了してください");
                Error.Exit();
                // Console.WriteLine("Excelが開いたままです。Excelを終了してください");
                // Console.WriteLine("このウィンドウを閉じるには、任意のキーを押してください...");
                // Console.ReadKey();
                // Environment.Exit(0);
            }

            worksheet = workbook.Worksheet("Sheet1");
            // lastRow = (int)worksheet.Cell("C1").GetValue<int>();
            lastRow = worksheet.LastRowUsed().RowNumber();
        }

        public string[] 読み取り()
        {
            string[] 英単語 = new string[lastRow];
            for (int i = 0; i < lastRow; i++)
            {
                英単語[i] = worksheet.Cell(i + 1, 1).Value.ToString();
                英単語[i] = カッコ内の文字を消す(英単語[i]);
            }
            return 英単語;
        }

        public void 書き込み(string[,] 発音記号, bool カッコを外す = false)
        {
            for (int i = 0; i < 発音記号.Length / 2; i++)
            {
                if (カッコを外す || 発音記号[0, i] == "Not found")
                {
                    worksheet.Cell(i + 1, 2).SetValue(発音記号[0, i]);
                }
                else
                {
                    worksheet.Cell(i + 1, 2).SetValue($"[{発音記号[0, i]}]");
                }
                worksheet.Cell(i + 1, 3).SetValue(発音記号[1, i]);
            }

            workbook.Save();
        }

        public void 起動()
        {
            string a = EXCEL実行ファイルのパス取得();
            Console.WriteLine(a);
            Process.Start(a, excelのパス);
        }

        private string カッコ内の文字を消す(string 単語)
        {
            Regex カッコの正規表現 = new(@"\(.+?\)");
            return カッコの正規表現.Replace(単語, "");
        }

        string EXCEL実行ファイルのパス取得()
        {
            // 操作するレジストリ・キーの名前
            string rKeyName = @"SOFTWARE\Microsoft\Windows\CurrentVersion\Extensions";
            // 取得処理を行う対象となるレジストリの値の名前
            string rGetValueName = "xlsx";

            // レジストリの取得
            try
            {
                // throw new NullReferenceException();
                // レジストリ・キーのパスを指定してレジストリを開く
                RegistryKey rKey = Registry.CurrentUser.OpenSubKey(rKeyName);

                // レジストリの値を取得
                string location = (string)rKey.GetValue(rGetValueName);

                // 開いたレジストリ・キーを閉じる
                rKey.Close();

                return location;
            }
            catch (NullReferenceException)
            {
                DirectoryInfo Excelexeのパスtxt = new(excelのパス);
                Excelexeのパスtxt = new(Excelexeのパスtxt.Parent.ToString() + @"\EXCEL.EXEのパス.txt");
                // Console.WriteLine(Excelexeのパスtxt);
                if (File.Exists(Excelexeのパスtxt.ToString()))
                {
                    return File.ReadLines(Excelexeのパスtxt.ToString()).ToString();
                }
                else
                {
                    using (var cofd = new CommonOpenFileDialog()
                    {
                        Title = "EXCEL.EXEを選択してください",
                        InitialDirectory = @"C:\\",
                        // フォルダ選択モードにする
                        IsFolderPicker = false,
                        AllowNonFileSystemItems = false,
                        Multiselect = false
                    })
                    {
                        if (cofd.ShowDialog() != CommonFileDialogResult.Ok)
                        {
                            return "";
                        }
                        using (FileStream fs = File.Create(Excelexeのパスtxt.ToString()))
                        {
                            byte[] info = new UTF8Encoding(true).GetBytes(cofd.FileName);
                            fs.Write(info, 0, info.Length);
                        }

                        return cofd.FileName;
                        // FileNameで選択されたフォルダを取得する
                        // System.Windows.MessageBox.Show($"{cofd.FileName}を選択しました");
                    }
                }
            }
        }
    }

    class Web操作
    {
        private int count;
        private string[] url = new string[0];
        const string weblio_url = @"https://ejje.weblio.jp/content/";
        private string[] html = new string[0];
        private Regex タグの正規表現 = new("<.+?>");
        public Web操作(string[] 英単語)
        {
            count = 英単語.Length;
            Array.Resize<string>(ref url, count);
            Array.Resize<string>(ref html, count);
            url = (Array.ConvertAll(英単語, conecturl));
        }

        private string conecturl(string input) { return $"{weblio_url}{input}"; }

        public string[,] 読み取り()
        {
            WebClient wc = new WebClient();
            string[,] 読み方と意味 = new string[2, count];
            var sw = new System.Diagnostics.Stopwatch();
            for (int i = 0; i < count; i++)
            {
                sw.Restart();
                while (true)
                {
                    try
                    {
                        // throw new System.Net.WebException();
                        html[i] = wc.DownloadString(url[i]);
                        break;
                    }
                    catch (System.Net.WebException)
                    {
                        Console.WriteLine("接続が切断されました。ネットワーク接続を確認してください。");
                        Console.WriteLine("任意のキーで再読み込み。");
                        Console.WriteLine("終了するには、Escキーを押してください...");
                        // Console.WriteLine(Console.ReadKey().Key.ToString());
                        // File.WriteAllText("./test.txt", Console.ReadKey(true).Key.ToString());
                        if (Console.ReadKey().Key.ToString().Equals("Escape"))
                            Environment.Exit(0);
                        Console.WriteLine("");
                    }
                }

                sw.Stop();
                Console.Write($"\"{url[i].Substring(31)}\"は{sw.ElapsedMilliseconds}ミリ秒でダウンロードされました");
                Console.WriteLine($" {i + 1}/{count}");
                string[] 分割 = html[i].Split("\n");
                int 検索行 = 0;
                try
                {
                    while (true)
                    {
                        int 行 = 分割[検索行].IndexOf(@"</span><span class=phoneticEjjeDc>(米国英語)");
                        if (行 != -1)
                        {
                            読み方と意味[0, i] = 分割[検索行].Substring(0, 行);
                            読み方と意味[0, i] = 読み方と意味[0, i].Substring(92);
                            読み方と意味[0, i] = タグの正規表現.Replace(読み方と意味[0, i], "");
                            // Console.WriteLine(読み方と意味[i]);
                            break;
                        }
                        検索行 += 1;
                    }
                }
                catch (System.IndexOutOfRangeException)
                {
                    Console.WriteLine($"\"{url[i].Substring(31)}\"の読み方は見つかりませんでした");
                    読み方と意味[0, i] = "Not found";
                }

                検索行 = 0;
                try
                {
                    while (true)
                    {
                        int 行 = 分割[検索行].IndexOf(@"主な意味");
                        if (行 != -1)
                        {
                            読み方と意味[1, i] = タグの正規表現.Replace(分割[検索行], "");
                            読み方と意味[1, i] = 読み方と意味[1, i].Substring(4);


                            break;
                        }
                        検索行 += 1;
                    }
                }
                catch (System.IndexOutOfRangeException)
                {
                    Console.WriteLine($"\"{url[i].Substring(31)}\"の意味は見つかりませんでした");
                    読み方と意味[1, i] = "Not found";
                }
                // Console.WriteLine(読み方と意味[1, i]);

            }
            return 読み方と意味;
        }
    }

    class Error
    {
        public static void Massage(string massage)
        {
            Console.WriteLine(massage);
            Console.WriteLine("このウィンドウを閉じるには、任意のキーを押してください...");
            Console.ReadKey();
        }

        public static void Exit()
        {
            Environment.Exit(0);
        }
    }
}