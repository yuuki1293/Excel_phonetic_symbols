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
        ex.シートを削除("Sheet2");
        ex.セーブ();
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
        worksheet = workbook.Worksheet("Sheet1");
    }
    public void 書き込み(string セル, string 値)
    {
        worksheet.Cell(セル).Value = "Hello";
    }
    public void 書き込み(string セル, int 値)
    {
        worksheet.Cell(セル).SetValue(値);
    }

    public void セーブ()
    {
        workbook.Save();
    }
    public void セーブ(string 名前)
    {
        workbook.SaveAs(名前);
    }
    public void シートを削除(string 名前)
    {
        workbook.Worksheet(名前).Delete();
    }
}

class Web操作
{
    public Web操作()
    {

    }
}
