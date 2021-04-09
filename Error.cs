using System;

namespace 発音記号
{
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