using sqlToExcel;
using System;

class Program
{
    static void Main(string[] args)
    {
        string connectionString = "Data Source=localhost\\SQLEXPRESS;Initial Catalog=planergy;Integrated Security=True;";
        string sqlFilePath = @"C:\Users\claud\OneDrive\Desktop\planergy\Input\RAPPORTI_UTENZA.sql";
        string excelPath = @"C:\Users\claud\OneDrive\Desktop\planergy\Output\risultato.xlsx";

        try
        {
            var converter = new SqlToExcelConverter(connectionString);
            converter.ConvertSqlFileToExcel(sqlFilePath, excelPath);
            Console.WriteLine("Conversione completata con successo!");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Si è verificato un errore: {ex.Message}");
        }
    }
}