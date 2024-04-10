# TableToJsonlConverter

## Summary
以下の記事に影響を受けて作成しました。

[Qiita - Snowflake Snowpipeを本番導入する前に読むやつ](https://zenn.dev/pei0804/articles/snowflake-snowpipe-production-ready)

ローカルのデータをJson Linesに簡単に変換することを目的としています。現時点では以下に対応しています。

- Excel
- CSV（or TSV）
- Microsoft SQL Server
- SQLite

## Getting Started

[Nuget - TableToJsonlConverter](https://www.nuget.org/packages/TableToJsonlConverter/)

## Source Code

- Visual Studio 2022
- .NET 6.0

## How to Use

エクセルからJsonLinesに変換する

```
    string ifile = @"TestFiles\Excel\test_base.xlsx";   // input file path
    bool headerf = true;                                // Whether the header is present
    int scol = 1;                                       // Start column Numboer (Numbers starting with 1)
    int srow = 1;                                       // Start row Numboer (Numbers starting with 1)
    int chcol = 1;                                      // The column number that must contain a value (the import ends where this value is missing)
    int sheetno = 0;                                    // Excel Sheet Index (Numbers starting with 0)

    ZkExcelToJsonl test = new ZkExcelToJsonl(ifile, headerf, scol, srow, chcol, sheetno);   // Initialize
    test.Read();                                                                            // File Read
    var header = test.GetHeader()!;                                                          // Get header Information


    header.ForEach(col => { Console.WriteLine(col.Col + "->" + col.ColumnName); });
    // output
    // 1->header1
    // 2->header2
    // test         ←LineBreak
    // 3->header3
    // 4->header4


    Console.WriteLine(test.JsonLines);
    // output
    //{"header1": "https://www.premium-tsubu-hero.net/","header2\ntest": "Brow1","header3": "Cro    w1","header4": "Drow1"}
    // { "header1": "Arow2","header2\ntest": "Brow2","header3": "Crow2","header4": "A series of clinical trials have proven conclusively that the new medication is effective in treating the virus.\nBy the time she retired, the human resources director at Mycroft Enterprises had worked with the company for fifty years."}
    // { "header1": "Arow3","header2\ntest": "Brow3","header3": "Crow3","header4": "Dro\"w\"3"}
    // { "header1": "Arow4","header2\ntest": "Brow4","header3": "Cr,ow4","header4": "Drow4"}
    // { "header1": "Arow5","header2\ntest": "Brow5","header3": "Crow5","header4": "Drow5"}


    string ofile = @"TestFiles\test_base.xlsx.json";

    // output Json Lines File
    test.Write(ofile);

    // outpu Json Lines .gz
    test.CompressWrite(ofile);
```

Excel to Json Lines

```
var test = new ZkExcelToJsonl()
{
    InputPath = @"C:\test.xlsx",    // input file path for .xlsx
    HeaderF = true,                 // hedder -> true:exist false:not exist
    StartCol = 1,                   // start column -> 1~
    StartRow = 1,                   // start row -> 1~
    CheckCol = 1,                   // not null column -> 1~
    SheetNo = 0                     // excel sheet index -> 0~
};

test.Read();

string outpath = @"C:\test.json";
test.Write(outpath);                // json lines file
test.CompressWrite(outpath);        // compress json lines file .gz

```

Csv to Json Lines

```
var test = new ZkCsvToJsonl()
{
    InputPath = @"C:\test.csv",     // input file path for .csv
    Encoding = Encoding.UTF8,       // file encoding
    HeaderF = true,                 // hedder -> true:exist false:not exist
    Delimiter = ",",                // delimita caractor
};
test.Read();

string outpath = @"C:\test.json";
test.Write(outpath);                // json lines file
test.CompressWrite(outpath);        // compress json lines file .gz
```

SQLite to Json Lines

```
var test = new ZkSQLiteToJsonl()
{
    ConnectionString = @"Data Source=test.db",
    SQLCommand = "select * From [TEST]"
};
test.Read();

string outpath = @"C:\test.json";
test.Write(outpath);                // json lines file
test.CompressWrite(outpath);        // compress json lines file .gz
```

Microsoft SQL Server to Json Lines

```
var test = new ZkSQLServerToJsonl()
{
    ConnectionString = @"Data Source=PCNAME\SQLEXPRESS;Initial Catalog=AdventureWorks2019;User ID=testuser;Password=testuser;Encrypt=False;Trust Server Certificate=True",
    SQLCommand = "select top 100 * From Person.Person"
};

test.Read();

string outpath = @"C:\test.json";
test.Write(outpath);                // json lines file
test.CompressWrite(outpath);        // compress json lines file .gz
```



