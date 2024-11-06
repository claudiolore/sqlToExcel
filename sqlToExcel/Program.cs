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
        try
        {
            string sqlFilePath = @"C:\Users\claud\OneDrive\Desktop\planergy\sql\RAPPORTI_UTENZA.sql";
            string excelPath = @"C:\Users\claud\OneDrive\Desktop\testsqlexcel\risultato.xlsx";

            try
            {
 
                var converter = new SqlToExcelConverter();
                converter.ConvertSqlFileToExcel(sqlFilePath, excelPath);
                Console.WriteLine($"Conversione completata con successo!");
                Console.WriteLine($"File Excel creato in: {excelPath}");

                System.Diagnostics.Process.Start("explorer.exe", excelPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Si è verificato un errore: {ex.Message}");
            }
        }
        catch (FileNotFoundException ex)
        {
            Console.WriteLine($"File non trovato: {ex.Message}");
        }
        catch (IOException ex)
        {
            Console.WriteLine($"Errore di I/O: {ex.Message}");
        }
        catch (SQLiteException ex)
        {
            Console.WriteLine($"Errore database: {ex.Message}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Errore imprevisto: {ex.Message}");
            Console.WriteLine($"Stack Trace: {ex.StackTrace}");
        }
    }
}