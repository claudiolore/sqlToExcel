using System;
using System.Data.SQLite;
using ClosedXML.Excel;
using System.IO;
using System.Text.RegularExpressions;
using System.Collections.Generic;

namespace sqlToExcel
{
    public class SqlToExcelConverter
    {
        private SQLiteConnection InitializeInMemoryDatabase(string sqlFilePath)
        {
            var connection = new SqliteConnection("Data Source=:memory:");
            connection.Open();

            string sqlContent = File.ReadAllText(sqlFilePath);
            string[] sqlCommands = Regex.Split(sqlContent, @";(\r?\n|$)")
                                      .Where(cmd => !string.IsNullOrWhiteSpace(cmd))
                                      .ToArray();

            foreach (string command in sqlCommands)
            {
                if (!string.IsNullOrWhiteSpace(command))
                {
                    using (var cmd = new SQLiteCommand(command, connection))
                    {
                        try
                        {
                            cmd.ExecuteNonQuery();
                        }
                        catch (SQLiteException ex)
                        {
                            Console.WriteLine($"Errore nell'esecuzione del comando: {command}");
                            Console.WriteLine($"Errore: {ex.Message}");
                            throw;
                        }
                    }
                }
            }

            return connection;
        }

        private List<string> GetAllTables(SQLiteConnection connection)
        {
            List<string> tables = new List<string>();
            using (var cmd = new SQLiteCommand(
                "SELECT name FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%';",
                connection))
            {
                using (var reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        tables.Add(reader.GetString(0));
                    }
                }
            }
            return tables;
        }

        private void ExportTableToExcel(string tableName, SQLiteConnection connection, XLWorkbook workbook)
        {
            using (var command = new SQLiteCommand($"SELECT * FROM [{tableName}]", connection))
            using (var reader = command.ExecuteReader())
            {
                var worksheet = workbook.Worksheets.Add(tableName);

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

                // Auto-fit delle colonne
                worksheet.Columns().AdjustToContents();
            }
        }

        public void ConvertSqlFileToExcel(string sqlFilePath, string excelFilePath)
        {
            try
            {
                using (var connection = InitializeInMemoryDatabase(sqlFilePath))
                {
                    var tables = GetAllTables(connection);

                    if (tables.Count == 0)
                    {
                        throw new Exception("Nessuna tabella trovata nel database");
                    }

                    using (var workbook = new XLWorkbook())
                    {
                        foreach (var table in tables)
                        {
                            ExportTableToExcel(table, connection, workbook);
                            Console.WriteLine($"Tabella {table} esportata con successo");
                        }

                        workbook.SaveAs(excelFilePath);
                    }
                }
            }
            catch (FileNotFoundException)
            {
                Console.WriteLine($"File SQL non trovato: {sqlFilePath}");
                throw;
            }
            catch (SQLiteException ex)
            {
                Console.WriteLine($"Errore SQLite: {ex.Message}");
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
