using Microsoft.WindowsAPICodePack.Dialogs;
using System;
using System.IO;
using ClosedXML.Excel;
using System.Text.RegularExpressions;
using System.Diagnostics;
using Microsoft.Win32;
using System.Text;

namespace 発音記号
{
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
            // Console.WriteLine(a);
            excelのパス = $"\"{excelのパス}\"";
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
}