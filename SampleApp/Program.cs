// See https://aka.ms/new-console-template for more information
using DocumentFormat.OpenXml.Spreadsheet;
using System.Diagnostics;
using TableToJsonlConverter.Conveters;
Console.WriteLine("Hello, World!");


void FromExcel()
{
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
}

FromExcel();