using System;
using System.Net;
using System.Text.RegularExpressions;

namespace 発音記号
{
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
}