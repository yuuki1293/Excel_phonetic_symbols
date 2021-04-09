using System;
using System.IO;

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
}