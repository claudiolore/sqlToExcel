using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ClosedXML.Excel;
using System.IO;
using System.Data.SqlClient;

namespace sqlToExcel
{

    public class SqlToExcelConverter
    {
        private readonly string _connectionString;

        public SqlToExcelConverter(string connectionString)
        {
            _connectionString = connectionString;
        }

        public void ConvertSqlFileToExcel(string sqlFilePath, string excelFilePath)
        {
            try
            {
                // Legge il contenuto del file SQL
                string sqlQuery = File.ReadAllText(sqlFilePath);

                using (var connection = new SqlConnection(_connectionString))
                {
                    connection.Open();
                    using (var command = new SqlCommand(sqlQuery, connection))
                    using (var reader = command.ExecuteReader())
                    {
                        using (var workbook = new XLWorkbook())
                        {
                            var worksheet = workbook.Worksheets.Add("Risultati");

                            // Aggiunge intestazioni
                            for (int i = 0; i < reader.FieldCount; i++)
                            {
                                worksheet.Cell(1, i + 1).Value = reader.GetName(i);
                            }

                            // Aggiunge dati
                            int row = 2;
                            while (reader.Read())
                            {
                                for (int i = 0; i < reader.FieldCount; i++)
                                {
                                    worksheet.Cell(row, i + 1).Value = reader[i]?.ToString() ?? "";
                                }
                                row++;
                            }

                            workbook.SaveAs(excelFilePath);
                        }
                    }
                }
            }
            catch (FileNotFoundException)
            {
                Console.WriteLine($"File SQL non trovato: {sqlFilePath}");
                throw;
            }
            catch (SqlException ex)
            {
                Console.WriteLine($"Errore SQL: {ex.Message}");
                throw;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Errore generico: {ex.Message}");
                throw;
            }
        }
    }
}
