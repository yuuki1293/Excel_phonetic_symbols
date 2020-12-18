using System;
using System.IO;
using ClosedXML.Excel;

class 発音記号
{
    static void Main(string[] args)
    {
        string exeファイルのパス = AppDomain.CurrentDomain.BaseDirectory.TrimEnd('\\');
        DirectoryInfo 作業ディレクトリ = new DirectoryInfo(exeファイルのパス);
        さかのぼる(ref 作業ディレクトリ, 3);
        Excel操作 ex = new Excel操作(作業ディレクトリ);
        ex.書き込み("A1", "Hello World");
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
    public Excel操作(DirectoryInfo 作業ディレクトリ)
    {
        excelのパス = 作業ディレクトリ.ToString() + @"\発音記号.xlsx";
        workbook = new XLWorkbook(excelのパス);
        worksheet = workbook.Worksheets.Add("Sheet1");

        Console.WriteLine(workbook.Worksheet("Sheet1"));
    }
    public void 書き込み(string セル, string 値)
    {
        worksheet.Cell(セル).Value = "Hello";
    }
    public void 書き込み(string セル, int 値)
    {
        worksheet.Cell(セル).SetValue(値);
    }
    // public void 書き込み(int 行, int 列, string 値)
    // {
    //     int 桁数 = (int)Math.Ceiling(Math.Log10(列));
    //     string アルファベット = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
    //     string セル =
    //     for (int i = 0; i < 桁数; i++)
    //     {
    //         アルファベット[列 % 26];
    //     }
    // }
}

class Web操作
{
    public Web操作()
    {

    }
}
