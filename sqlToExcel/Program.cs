using System;
using System.Data.SQLite;
using ClosedXML.Excel;
using System.IO;
using System.Text.RegularExpressions;
using System.Collections.Generic;
using sqlToExcel;

class Program
{
    static void Main(string[] args)
    {
        string sqlFilePath = @"C:\Users\claud\OneDrive\Desktop\planergy\Input\RAPPORTI_UTENZA.sql";
        string excelPath = @"C:\Users\claud\OneDrive\Desktop\planergy\Output\risultato.xlsx";

        try
        {
            var converter = new SqlToExcelConverter();
            converter.ConvertSqlFileToExcel(sqlFilePath, excelPath);
            Console.WriteLine("Conversione completata con successo!");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Si è verificato un errore: {ex.Message}");
        }
    }
}